"""Behavioural tests for persisted monthly billing and explicit sparse daily load evidence."""
import copy
import json
import sqlite3
from contextlib import closing
import tempfile
from sqlite_tempdir import SqliteTemporaryDirectory
import unittest
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote, urlencode

from fixtures import usage
from redink_server.config import Settings
from redink_server.protocol import normalize_event, parse_legacy_path, parse_legacy_query, ProtocolError
from redink_server.store import Store, iso


class StoreTests(unittest.TestCase):
    def setUp(self):
        self.temp = SqliteTemporaryDirectory()
        self.directory = Path(self.temp.name)
        self.now = datetime(2026, 10, 30, 12, 0, tzinfo=timezone.utc)
        self.settings = Settings(data_dir=self.directory, credentials={}, recent_report_window_seconds=3600)
        self.store = Store(self.settings, clock=lambda: self.now)

    def tearDown(self):
        self.temp.cleanup()

    def ingest(self, *, direct=False, **kwargs):
        return self.store.ingest(normalize_event(usage(**kwargs)), live_request=direct)

    def daily(self, day="2026-10-30"):
        return self.store.dashboard(day[:7], day)["day"]["users"]

    def fresh(self, **kwargs):
        return usage(last_use=iso(self.now - timedelta(seconds=10)), reported=iso(self.now), interval_minutes=5, **kwargs)

    def test_monthly_late_reporting_stays_in_original_month(self):
        self.ingest(month="2026-08")
        self.ingest(month="2026-09")
        rows = self.store.monthly_rows()
        self.assertEqual([(r["month"], r["users"]) for r in rows], [("2026-08", 1), ("2026-09", 1)])
        dashboard = self.store.dashboard("2026-10", "2026-10-30")
        self.assertEqual(dashboard["month"]["users"], 0)
        self.assertEqual(dashboard["day"]["users"], 0)

    def test_four_legacy_routes_deduplicate_and_enrich_product(self):
        full = usage(product="Red Ink Pro")
        short = {"ri_lc_v": 1, "t": "U", "m": full["month"], "n": full["networkKey"], "ns": full["networkScope"], "p": full["productIdSafe"], "h": full["hostId"], "u": full["userKey"], "b": full["billingEventId"], "e": full["hostEventId"], "a": 1}
        path = "/collect/RI_LC_v1/" + "/".join(key + "/" + quote(str(short[key]), safe="") for key in ("m", "n", "ns", "p", "h", "u", "b", "e", "a")) + "/r/0123456789abcdef"
        self.store.ingest(parse_legacy_path(path))
        self.store.ingest(parse_legacy_query(urlencode(short)))
        short["p"] = full["productId"]
        self.store.ingest(normalize_event(short))
        self.store.ingest(normalize_event(full))
        rows = self.store.monthly_rows()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["product_id"], "Red Ink Pro")
        self.assertEqual(rows[0]["users"], 1)
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["month"]["events"], 1)

    def test_one_user_across_three_hosts_is_one_billable_user(self):
        for host in ("WD", "EX", "OL"):
            self.ingest(host=host)
        self.ingest(host="WD", user="bob")
        rows = self.store.monthly_rows()
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["users"], 2)
        self.assertEqual((row["word_users"], row["excel_users"], row["outlook_users"]), (2, 1, 1))
        self.assertEqual(len(self.store.user_rows("2026-10")), 2)

    def test_scopes_and_products_are_separate_billing_groups(self):
        for scope, product in (("ONE.EXAMPLE", "PRO"), ("ONE.EXAMPLE", "TEAM"), ("TWO.EXAMPLE", "PRO")):
            self.ingest(scope=scope, product=product)
        rows = self.store.monthly_rows()
        self.assertEqual(len(rows), 3)
        self.assertTrue(all(r["users"] == 1 for r in rows))

    def test_legacy_timestamps_do_not_invent_daily_evidence(self):
        event = normalize_event(self.fresh())
        for live in (False, True):
            self.store.ingest(event, live_request=live)
        self.assertEqual(self.daily(), 0)
        row = self.store.monthly_rows()[0]
        self.assertEqual((row["users"], row["monthly_only_users"], row["observed_user_days"]), (1, 1, 0))

    def test_daily_evidence_survives_closure_and_receipt_window(self):
        self.ingest(loaded_dates=["2026-10-30"])
        self.now += timedelta(hours=5)
        self.assertEqual(self.daily(), 1)
        self.now += timedelta(days=1)
        self.assertEqual(self.daily("2026-10-30"), 1)
        self.assertEqual(self.daily("2026-10-31"), 0)
        self.assertEqual(self.store.monthly_rows()[0]["users"], 1)

    def test_sparse_offline_days_are_backfilled_without_filling_gaps(self):
        self.ingest(month="2026-09", loaded_dates=["2026-09-01", "2026-09-03", "2026-09-30"])
        self.assertEqual(self.daily("2026-09-01"), 1)
        self.assertEqual(self.daily("2026-09-02"), 0)
        self.assertEqual(self.daily(), 0)
        self.assertEqual(self.store.monthly_rows()[0]["observed_user_days"], 3)
        self.assertEqual(self.store.monthly_rows()[0]["users"], 1)

    def test_same_user_multiple_hosts_devices_and_replays_count_once_per_day(self):
        for host in ("WD", "EX", "OL", "WD"):
            self.ingest(host=host, loaded_dates=["2026-10-29", "2026-10-30"])
        self.assertEqual(self.daily(), 1)
        row = self.store.monthly_rows()[0]
        self.assertEqual((row["users"], row["observed_user_days"]), (1, 2))
        self.assertEqual((row["word_users"], row["excel_users"], row["outlook_users"]), (1, 1, 1))

    def test_aggregate_hosts_need_no_synthetic_host_event_ids(self):
        self.ingest(loaded_dates=["2026-10-30"], loaded_hosts=["WD", "EX", "OL"])
        row = self.store.monthly_rows()[0]
        self.assertEqual((row["word_users"], row["excel_users"], row["outlook_users"]), (1, 1, 1))
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["month"]["events"], 1)
        self.assertEqual(self.store.user_rows("2026-10")[0]["hosts"], "EX,OL,WD")

    def test_late_missing_date_revises_evidence_but_not_monthly_units(self):
        self.ingest(loaded_dates=["2026-10-30"])
        self.assertEqual(self.store.monthly_rows()[0]["revision"], 1)
        self.ingest(loaded_dates=["2026-10-01", "2026-10-30"])
        revised = self.store.monthly_rows()[0]
        self.assertEqual((revised["users"], revised["observed_user_days"], revised["revision"]), (1, 2, 2))
        self.ingest(loaded_dates=["2026-10-30", "2026-10-01"])
        self.assertEqual(self.store.monthly_rows()[0]["revision"], 2)
        self.assertEqual(self.store.user_rows("2026-10")[0]["loaded_dates"], "2026-10-01,2026-10-30")

    def test_daily_evidence_is_valid_from_files_and_http(self):
        for direct in (False, True):
            self.ingest(direct=direct, loaded_dates=["2026-10-30"])
        self.assertEqual(self.daily(), 1)
        self.assertEqual(self.store.monthly_rows()[0]["observed_user_days"], 1)

    def test_client_calendar_at_utc_month_boundary(self):
        self.now = datetime(2026, 10, 31, 23, 30, tzinfo=timezone.utc)
        self.ingest(month="2026-11", loaded_dates=["2026-11-01"])
        self.assertEqual(self.daily("2026-11-01"), 1)
        self.assertEqual(self.daily("2026-10-31"), 0)

    def test_future_daily_evidence_rolls_back_monthly_identity(self):
        with self.assertRaises(ProtocolError):
            self.ingest(month="2026-11", loaded_dates=["2026-11-02"])
        self.assertEqual(self.store.monthly_rows(), [])
        self.assertEqual(self.store.check(), "ok")

    def test_invalid_future_month_rolls_back(self):
        with self.assertRaises(ProtocolError): self.ingest(month="2027-01")
        self.assertEqual(self.store.monthly_rows(), [])
        self.assertEqual(self.store.check(), "ok")

    def test_identity_conflict_is_transactional(self):
        original = normalize_event(usage())
        self.store.ingest(original)
        conflict = normalize_event(usage(user="bob"))
        conflict["billing_event_id"] = original["billing_event_id"]
        before = self.store.monthly_rows()
        with self.assertRaises(ProtocolError): self.store.ingest(conflict)
        self.assertEqual(before, self.store.monthly_rows())
        self.assertEqual(len(self.store.user_rows("2026-10")), 1)
        conflict = normalize_event(usage(user="charlie"))
        conflict["host_event_id"] = original["host_event_id"]
        with self.assertRaises(ProtocolError): self.store.ingest(conflict)
        self.assertEqual(before, self.store.monthly_rows())

    def test_twenty_concurrent_connections_deduplicate_atomically(self):
        event = normalize_event(usage(loaded_dates=["2026-10-30"]))
        with ThreadPoolExecutor(max_workers=20) as executor:
            results = list(executor.map(lambda _: self.store.ingest(event), range(20)))
        self.assertEqual(sum(not r["duplicate"] for r in results), 1)
        self.assertEqual(self.store.monthly_rows()[0]["users"], 1)
        self.assertEqual(self.store.monthly_rows()[0]["revision"], 1)
        self.assertEqual(self.daily(), 1)
        self.assertEqual(self.store.check(), "ok")

    def test_restart_and_verified_backup_preserve_history(self):
        self.ingest(month="2026-03")
        self.ingest(month="2026-10")
        before = self.store.monthly_rows()
        reopened = Store(self.settings, clock=lambda: self.now)
        self.assertEqual(reopened.monthly_rows(), before)
        destination = self.directory / "backups" / "backup.sqlite3"
        reopened.backup(destination)
        with closing(sqlite3.connect(destination)) as connection, connection:
            self.assertEqual(connection.execute("PRAGMA quick_check").fetchone()[0], "ok")
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM billing_usage").fetchone()[0], 2)
        with self.assertRaises(ValueError): reopened.backup(destination)
        with self.assertRaises(ValueError): reopened.backup(self.settings.database_path)

    def test_error_events_are_deduplicated_without_adding_usage(self):
        event = normalize_event({"v": 1, "t": "E", "errorEventId": "x" + "a"*32, "errorCode": "local_marker_missing", "relatedMonth": "2026-09"})
        self.assertFalse(self.store.ingest(event)["duplicate"])
        self.assertTrue(self.store.ingest(event)["duplicate"])
        self.assertEqual(self.store.monthly_rows(), [])
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["health"]["pending_errors"], 1)

    def test_anonymous_replay_erases_reversible_identity(self):
        wire = usage(); wire.update(anon=False, userObf="o1_YWxpY2U")
        event = normalize_event(wire)
        self.store.ingest(event)
        self.store.ingest(normalize_event(usage()))
        self.store.ingest(event)
        with self.store.connect() as connection:
            row = connection.execute("SELECT anonymous,user_obf FROM billing_usage").fetchone()
        self.assertEqual(tuple(row), (1, ""))

    def test_report_snapshots_immutable_and_content_addressed(self):
        report = {"month": "2026-10", "users": 5}
        one = self.store.save_report_snapshot(report)
        self.now += timedelta(seconds=30)
        two = self.store.save_report_snapshot(report)
        self.assertEqual(one["report_id"], two["report_id"])
        self.assertEqual(one["sha256"], two["sha256"])
        with self.store.connect() as connection:
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM report_snapshots").fetchone()[0], 1)


if __name__ == "__main__": unittest.main()
