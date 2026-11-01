import csv
import io
import json
import tempfile
import unittest
from dataclasses import replace
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path

from fixtures import usage
from redink_server.config import ConfigError, initialize, load_settings
from redink_server.protocol import normalize_event
from redink_server.reports import csv_text, monthly_report, report_csv
from redink_server.store import Store


class ReportConfigTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.path = Path(self.temporary.name)
        self.credentials = initialize(self.path)
        self.settings = load_settings(self.path)

    def test_init_never_overwrites_and_does_not_store_clear_credentials(self):
        original = (self.path / "secrets.json").read_text()
        with self.assertRaises(ConfigError):
            initialize(self.path)
        for value in self.credentials.values():
            self.assertNotIn(value, original)
        self.assertEqual(original, (self.path / "secrets.json").read_text())

    def test_unknown_or_mistyped_options_fail(self):
        config = self.path / "server.toml"
        original = config.read_text()
        for content in (original.replace("port = 8765", 'port = "8765"'), original + '\n[misspelled]\na=1\n', original.replace('unit_price = ""', 'unit_price = "NaN"')):
            config.write_text(content)
            with self.assertRaises(ConfigError):
                load_settings(self.path)
        config.write_text(original)

    def test_csv_formula_injection_and_quoting(self):
        text = csv_text([{"name": '=HYPERLINK("bad")', "id": " \t+1", "safe": "alpha, beta", "number": 2}], ["name", "id", "safe", "number"])
        row = next(csv.DictReader(io.StringIO(text.lstrip("\ufeff"))))
        self.assertTrue(row["name"].startswith("'="))
        self.assertTrue(row["id"].startswith("'"))
        self.assertEqual(row["safe"], "alpha, beta")
        self.assertEqual(row["number"], "2")

    def test_optional_decimal_prices_are_exact_and_host_totals_not_billed(self):
        settings = replace(self.settings, unit_price="19.95", currency="CHF")
        store = Store(settings, clock=lambda: datetime(2026, 10, 30, tzinfo=timezone.utc))
        for host in ("WD", "EX", "OL"):
            store.ingest(normalize_event(usage(host=host)))
        store.ingest(normalize_event(usage(user="bob")))
        report = monthly_report(store, "2026-10", "2026-10")
        self.assertEqual(report["units"], 2)
        self.assertEqual(report["amount"], "39.90")
        self.assertEqual(report["rows"][0]["word_users"], 2)
        self.assertIn("content_sha256", report_csv(report))
        self.assertEqual(len(report["content_sha256"]), 64)

    def test_empty_month_is_no_evidence_and_late_report_revises_history(self):
        store = Store(self.settings, clock=lambda: datetime(2026, 10, 30, tzinfo=timezone.utc))
        report = monthly_report(store, "2026-09", "2026-10")
        self.assertEqual(report["rows"], [])
        self.assertIn("not confirmed zero", report["coverage"])
        store.ingest(normalize_event(usage(month="2026-09")))
        after = monthly_report(store, "2026-09", "2026-10")
        self.assertEqual(after["rows"][0]["month"], "2026-09")
        self.assertNotIn("amount", after)


if __name__ == "__main__":
    unittest.main()
