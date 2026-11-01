import json
import re
import sqlite3
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch
from urllib.parse import urlencode

from werkzeug.security import generate_password_hash

from fixtures import usage
from redink_server.config import Settings, token_hash
from redink_server.store import Store
from redink_server.web import create_app


class WebTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.ingest_token = "test-ingestion-" + "a" * 32
        self.config_token = "test-configuration-" + "b" * 32
        credentials = {"session_secret": "c" * 64,
                       "admin_password_hash": generate_password_hash("correct-test-password", method="pbkdf2:sha256:1000"),
                       "ingestion_token_hash": token_hash(self.ingest_token), "config_token_hash": token_hash(self.config_token)}
        self.settings = Settings(self.root, credentials)
        self.settings.config_directory.mkdir()
        self.now = datetime(2026, 10, 30, 10, 0, tzinfo=timezone.utc)
        self.store = Store(self.settings, clock=lambda: self.now)
        self.app = create_app(self.settings, self.store)
        self.app.testing = True
        self.client = self.app.test_client()
        self.endpoint = "/collect/" + self.ingest_token

    def signin(self):
        response = self.client.get("/login")
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', response.text).group(1)
        return self.client.post("/login", data={"csrf_token": csrf, "password": "correct-test-password"})

    def test_dashboard_and_exports_require_admin(self):
        self.assertEqual(self.client.get("/").status_code, 302)
        self.assertEqual(self.client.get("/api/dashboard").status_code, 401)
        self.assertEqual(self.client.get("/api/reports/monthly.csv").status_code, 401)
        self.assertEqual(self.signin().status_code, 302)
        self.assertEqual(self.client.get("/").status_code, 200)
        response = self.client.get("/api/dashboard")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["month"]["users"], 0)

    def test_login_csrf_and_throttling(self):
        self.assertEqual(self.client.post("/login", data={"password": "correct-test-password"}).status_code, 400)
        self.assertEqual(self.client.post("/login", data={"csrf_token": "ä", "password": "correct-test-password"}).status_code, 400)
        self.client.get("/login")
        with self.client.session_transaction() as session:
            csrf = session["csrf"]
        for _ in range(5):
            self.assertEqual(self.client.post("/login", data={"csrf_token": csrf, "password": "wrong"}).status_code, 401)
        self.assertEqual(self.client.post("/login", data={"csrf_token": csrf, "password": "wrong"}).status_code, 429)

    def test_authentication_is_separate_for_configuration_and_collection(self):
        self.assertEqual(self.client.post("/collect/wrong", json=usage()).status_code, 401)
        self.assertEqual(self.client.post("/collect/" + self.config_token, json=usage()).status_code, 401)
        self.assertEqual(self.client.post("/api/collect", json=usage(), headers={"Authorization": "Bearer " + self.ingest_token}).status_code, 200)
        self.assertEqual(self.client.post("/api/collect", json=usage(), headers={"Authorization": "Bearer " + self.config_token}).status_code, 401)

    def test_all_original_web_transports_ack_once(self):
        wire = usage(product="Red Ink Pro")
        short = {"ri_lc_v": 1, "t": "U", "m": wire["month"], "n": wire["networkKey"], "ns": wire["networkScope"],
                 "p": wire["productIdSafe"], "h": wire["hostId"], "u": wire["userKey"], "b": wire["billingEventId"],
                 "e": wire["hostEventId"], "a": 1, "r": "abc"}
        path = self.endpoint + "/RI_LC_v1/" + "/".join(str(v) for k in ("m", "n", "ns", "p", "h", "u", "b", "e", "a", "r") for v in (k, short[k]))
        replies = [self.client.get(path), self.client.get(self.endpoint + "?" + urlencode(short)),
                   self.client.post(self.endpoint, json=wire), self.client.post(self.endpoint, data=short)]
        self.assertTrue(all(reply.status_code == 200 and reply.json["accepted"] for reply in replies), [r.text for r in replies])
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["month"]["users"], 1)
        self.assertEqual(self.store.monthly_rows()[0]["product_id"], "Red Ink Pro")
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["day"]["users"], 0)

    def test_fresh_repeated_monthly_marker_does_not_duplicate(self):
        wire = usage(last_use="2026-10-30T10:00:00Z", reported="2026-10-30T10:00:00Z", interval_minutes=360, loaded_dates=["2026-10-30"])
        for _ in range(3):
            self.assertEqual(self.client.post(self.endpoint, json=wire).status_code, 200)
        dashboard = self.store.dashboard("2026-10", "2026-10-30")
        self.assertEqual(dashboard["month"]["users"], 1)
        self.assertEqual(dashboard["day"]["users"], 1)
        self.assertEqual(dashboard["month"]["observed_user_days"], 1)

    def test_daily_form_http_and_authenticated_exports_keep_original_dates(self):
        wire = usage(loaded_dates="2026-10-01,2026-10-30", loaded_hosts="WD,OL")
        for _ in range(2):
            response = self.client.post(self.endpoint, data=wire)
            self.assertEqual(response.status_code, 200)
            self.assertTrue(response.json["accepted"])
        self.assertEqual(self.client.get("/api/reports/daily.csv?month=2026-10").status_code, 401)
        self.signin()
        data = self.client.get("/api/dashboard?month=2026-10&day=2026-10-01").json
        self.assertEqual(data["day"]["users"], 1)
        self.assertEqual(data["month"]["users"], 1)
        self.assertEqual(data["month"]["observed_user_days"], 2)
        self.assertEqual(self.client.get("/api/dashboard?day=2026-02-30").status_code, 400)
        response = self.client.get("/api/reports/daily.csv?month=2026-10")
        self.assertEqual(response.status_code, 200)
        self.assertIn("2026-10-01", response.text)
        self.assertIn("2026-10-30", response.text)
        self.assertNotIn("2026-10-02", response.text)
        self.assertEqual(self.client.get("/api/reports/daily.csv?month=invalid").status_code, 400)

    def test_invalid_partial_duplicate_or_oversized_input_not_acknowledged(self):
        for raw in ("null", "[]", '{"v":1,"v":1}', '{"month":"2026-10"}', '{broken'):
            response = self.client.post(self.endpoint, data=raw, content_type="application/json")
            self.assertEqual(response.status_code, 400)
            self.assertFalse(response.json["accepted"])
        response = self.client.post(self.endpoint, data="x" * 70000, content_type="application/json")
        self.assertEqual(response.status_code, 413)
        self.assertFalse(response.json["accepted"])
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["month"]["users"], 0)

    def test_deeply_nested_json_is_rejected_without_logging_token(self):
        response = self.client.post(self.endpoint, data='[' * 2000 + '0' + ']' * 2000, content_type="application/json")
        self.assertEqual(response.status_code, 400)
        self.assertNotIn(self.ingest_token, response.text)

    def test_unpaired_unicode_and_astral_product_contract(self):
        from redink_server.protocol import safe_product_id
        self.assertEqual(safe_product_id("Red Ink 😀"), "Red_Ink___")
        wire = usage(product="Red Ink 😀")
        response = self.client.post(self.endpoint, json=wire)
        self.assertEqual(response.status_code, 200)
        wire["productId"] = "invalid\ud800"
        response = self.client.post(self.endpoint, data=json.dumps(wire), content_type="application/json")
        self.assertEqual(response.status_code, 400)

    def test_storage_failure_has_no_success_ack(self):
        with patch.object(self.store, "ingest", side_effect=sqlite3.OperationalError("test lock")):
            response = self.client.post(self.endpoint, json=usage())
        self.assertEqual(response.status_code, 503)
        self.assertFalse(response.json["accepted"])

    def test_head_cannot_collect(self):
        response = self.client.head(self.endpoint + "?" + urlencode(usage()))
        self.assertEqual(response.status_code, 405)
        self.assertEqual(self.store.monthly_rows(), [])

    def test_served_configuration_bytes_and_signature_are_unchanged(self):
        content = b"\xef\xbb\xbf; signed bytes\r\nLicenseCounterIntervall=360\r\n"
        (self.settings.config_directory / "redink.ini").write_bytes(content)
        (self.settings.config_directory / "redink.ini.sig").write_bytes(b"base64-signature\r\n")
        for name in ("redink.ini", "redink.ini.sig"):
            response = self.client.get("/config/" + self.config_token + "/" + name)
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.data, (self.settings.config_directory / name).read_bytes())
            response.close()
        self.assertEqual(self.client.get("/config/" + self.ingest_token + "/redink.ini").status_code, 401)
        self.assertEqual(self.client.get("/config/" + self.config_token + "/secrets.json").status_code, 404)
        self.assertEqual(self.client.get("/config/" + self.config_token + "/../secrets.json").status_code, 404)

    def test_config_symlink_cannot_expose_another_file(self):
        target = self.root / "private.ini"
        target.write_text("secret")
        try:
            (self.settings.config_directory / "redink.ini").symlink_to(target)
        except (OSError, NotImplementedError):
            self.skipTest("Symlinks unavailable")
        self.assertEqual(self.client.get("/config/" + self.config_token + "/redink.ini").status_code, 404)

    def test_security_headers_host_validation_and_session_logout(self):
        response = self.client.get("/healthz")
        self.assertIn("script-src 'self'", response.headers["Content-Security-Policy"])
        self.assertEqual(response.headers["X-Frame-Options"], "DENY")
        self.assertEqual(self.client.get("/healthz", headers={"Host": "evil.invalid"}).status_code, 400)
        self.signin()
        self.assertEqual(self.client.post("/logout").status_code, 400)
        self.client.get("/")
        with self.client.session_transaction() as session:
            csrf = session["csrf"]
        self.assertEqual(self.client.post("/logout", data={"csrf_token": csrf}).status_code, 302)
        self.assertEqual(self.client.get("/api/dashboard").status_code, 401)

    def test_report_ranges_and_snapshot_csrf(self):
        self.signin()
        self.client.post(self.endpoint, json=usage())
        self.assertEqual(self.client.get("/api/reports/monthly.csv?from=2026-11&to=2026-10").status_code, 400)
        self.assertEqual(self.client.get("/api/dashboard?month=invalid").status_code, 400)
        self.assertEqual(self.client.get("/api/reports/monthly.json?from=2026-10&to=2026-10").json["units"], 1)
        self.assertEqual(self.client.post("/api/reports/snapshot?from=2026-10&to=2026-10").status_code, 400)
        self.client.get("/")
        with self.client.session_transaction() as session:
            csrf = session["csrf"]
        response = self.client.post("/api/reports/snapshot?from=2026-10&to=2026-10", headers={"X-CSRF-Token": csrf})
        self.assertEqual(response.status_code, 201)
        self.assertTrue(response.json["report_id"].startswith("RI-"))

    def test_print_report_and_files_are_escaped(self):
        self.signin()
        self.client.post(self.endpoint, json=usage(product="<script>alert(1)</script>"))
        response = self.client.get("/reports/print?from=2026-10&to=2026-10")
        self.assertEqual(response.status_code, 200)
        self.assertNotIn("<script>alert", response.text)
        self.assertIn("&lt;script&gt;", response.text)


if __name__ == "__main__":
    unittest.main()
