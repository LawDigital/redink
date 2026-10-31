"""Daily evidence, migration and export tests at the public protocol boundary."""
import json
import sqlite3
from contextlib import closing
import tempfile
from sqlite_tempdir import SqliteTemporaryDirectory
import unittest
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode

from fixtures import usage
from redink_server.config import Settings
from redink_server.protocol import ProtocolError, normalize_event, parse_legacy_query
from redink_server.reports import daily_csv, monthly_report, users_csv, valid_day
from redink_server.store import Store


class DailyContractTests(unittest.TestCase):
    def setUp(self):
        self.temp = SqliteTemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.settings = Settings(Path(self.temp.name), {})
        self.now = datetime(2026, 10, 30, 12, tzinfo=timezone.utc)

    def test_json_query_form_dates_and_hosts_are_identical(self):
        wire = usage(loaded_dates=["2026-10-30", "2026-10-01"], loaded_hosts=["OL", "EX", "WD"])
        canonical = normalize_event(wire)
        form = dict(ri_lc_v=1, **wire)
        form.update(loadedDates="2026-10-30,2026-10-01", loadedHosts="OL,EX,WD")
        form.pop("v")
        query = parse_legacy_query(urlencode(form, doseq=True))
        self.assertEqual(canonical, query)
        form_arrays = {key: [value] if not isinstance(value, list) else value for key, value in form.items()}
        self.assertEqual(canonical, normalize_event(form_arrays))

    def test_sparse_dates_are_real_bounded_and_month_local(self):
        for dates in (["2026-10-00"], ["2026-10-32"], ["2026-09-30"], ["2026-10-1"], [None], [["2026-10-01"]], ["2026-10-01"] * 32):
            with self.subTest(dates=dates), self.assertRaises(ProtocolError):
                normalize_event(usage(loaded_dates=dates))
        with self.assertRaises(ProtocolError):
            normalize_event(usage(month="2026-02", loaded_dates=["2026-02-29"]))
        self.assertEqual(normalize_event(usage(month="2024-02", loaded_dates=["2024-02-29"]))["loaded_dates"], ["2024-02-29"])

    def test_alias_conflicts_and_host_injection_rejected(self):
        wire = usage(loaded_dates=["2026-10-30"])
        wire["loaded_dates"] = ["2026-10-29"]
        with self.assertRaises(ProtocolError): normalize_event(wire)
        with self.assertRaises(ProtocolError): normalize_event(usage(loaded_hosts=["EX", "<script>"]))
        self.assertEqual(normalize_event(usage(loaded_hosts=["excel"]))["loaded_hosts"], ["EX", "WD"])

    def test_daily_exports_do_not_change_monthly_billing_units(self):
        store = Store(self.settings, clock=lambda: self.now)
        store.ingest(normalize_event(usage(loaded_dates=["2026-10-01", "2026-10-30"])))
        store.ingest(normalize_event(usage(host="OL", loaded_dates=["2026-10-30"])))
        store.ingest(normalize_event(usage(user="legacy")))
        report = monthly_report(store, "2026-10", "2026-10")
        self.assertEqual(report["units"], 2)
        self.assertEqual(report["rows"][0]["observed_user_days"], 2)
        self.assertEqual(report["rows"][0]["monthly_only_users"], 1)
        csv = daily_csv(store, "2026-10")
        self.assertIn("2026-10-01", csv)
        self.assertNotIn("2026-10-02", csv)
        self.assertIn("loaded_days,loaded_dates", users_csv(store, "2026-10"))
        self.assertIn('"2026-10-01,2026-10-30"', users_csv(store, "2026-10"))

    def test_day_input_rejects_impossible_calendar_values(self):
        for value in ("2026-02-29", "2026-13-01", "2026-1-01", "", None):
            with self.subTest(value=value), self.assertRaises(ValueError): valid_day(value)
        self.assertEqual(valid_day("2024-02-29"), "2024-02-29")

    def test_genuine_v1_schema_migration_keeps_history_and_snapshot_without_invented_days(self):
        # This SQL is the exact SCHEMA constant from the delivered 1.0.0 source.
        schema = (Path(__file__).parent / "data" / "schema-v1.sql").read_text()
        wire = normalize_event(usage())
        stamp = "2026-10-30T12:00:00.000000Z"
        old_snapshot = '{"original_report":true,"units":1}'
        with closing(sqlite3.connect(self.settings.database_path)) as connection, connection:
            connection.executescript(schema)
            connection.execute("PRAGMA user_version=1")
            connection.execute("INSERT INTO billing_usage VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (
                wire["billing_event_id"], wire["month"], wire["network_key"], wire["network_scope"],
                json.dumps(wire["allowed_network_ids"]), wire["product_id"], wire["product_id_safe"], wire["user_key"],
                1, "", wire["first_use_utc"], stamp, stamp, "legacy import"))
            connection.execute("INSERT INTO host_usage VALUES (?,?,?,?,?,?)", (wire["host_event_id"], wire["billing_event_id"], "WD", stamp, stamp, stamp))
            connection.execute("INSERT INTO recent_reports VALUES (?,?,?,?,?)", (wire["host_event_id"], stamp, stamp, stamp, 360))
            connection.execute("INSERT INTO month_revisions VALUES (?,?,?)", ("2026-10", 7, stamp))
            connection.execute("INSERT INTO report_snapshots VALUES (?,?,?,?)", ("original-id", stamp, old_snapshot, "original-hash"))
        store = Store(self.settings, clock=lambda: self.now)
        self.assertEqual(store.monthly_rows()[0]["users"], 1)
        self.assertEqual(store.monthly_rows()[0]["word_users"], 1)
        self.assertEqual(store.monthly_rows()[0]["revision"], 7)
        self.assertEqual(store.daily_rows("2026-10-01", "2026-10-31"), [])
        with store.connect() as connection:
            self.assertEqual(connection.execute("PRAGMA user_version").fetchone()[0], 3)
            self.assertEqual(connection.execute("SELECT content_json FROM report_snapshots").fetchone()[0], old_snapshot)
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM recent_reports").fetchone()[0], 1)
        store.ingest(normalize_event(usage(loaded_dates=["2026-10-30"])))
        self.assertEqual(store.monthly_rows()[0]["revision"], 8)
        self.assertEqual(store.monthly_rows()[0]["users"], 1)
        self.assertEqual(store.daily_rows("2026-10-30", "2026-10-30")[0]["users"], 1)
        self.assertEqual(Store(self.settings, clock=lambda: self.now).monthly_rows(), store.monthly_rows())


if __name__ == "__main__": unittest.main()
