"""Import regression tests: source loss, repeated scans, append tails and identity."""
import json
import shutil
import tempfile
import unittest
from dataclasses import replace
from datetime import datetime, timedelta, timezone
from pathlib import Path

from fixtures import usage, log_line, log_filename
from redink_server.config import Settings
from redink_server.scanner import Scanner
from redink_server.store import Store, iso


class ScannerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp.name)
        self.source = self.directory / "input"
        self.source.mkdir()
        self.now = datetime(2026, 10, 30, 12, tzinfo=timezone.utc)
        self.settings = Settings(data_dir=self.directory / "data", credentials={}, recent_report_window_seconds=3600)
        self.store = Store(self.settings, clock=lambda: self.now)
        self.scanner = Scanner(self.settings, self.store)

    def tearDown(self):
        self.scanner.stop()
        self.temp.cleanup()

    def marker(self, name="marker.json", **kwargs):
        path = self.source / name
        path.write_text(json.dumps(usage(**kwargs)), encoding="utf-8-sig")
        return path

    def log(self, *, directory=None, machine="PC1", text=None):
        folder = directory or self.source
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / log_filename(machine=machine)
        path.write_text(log_line() if text is None else text, encoding="utf-8-sig")
        return path

    def test_repeated_counter_scans_and_outage_retain_monthly_history(self):
        path = self.marker(month="2026-08")
        self.assertEqual(self.scanner.scan_path(self.source, "counter")["records"], 1)
        self.assertEqual(self.scanner.scan_path(self.source, "counter")["records"], 1)
        self.assertEqual(self.store.monthly_rows()[0]["users"], 1)
        path.unlink(); self.source.rmdir()
        result = self.scanner.scan_path(self.source, "counter")
        self.assertEqual(result["errors"], 1)
        self.assertIn("retained", result["detail"])
        self.assertEqual(self.store.monthly_rows()[0]["month"], "2026-08")
        self.assertEqual(self.store.monthly_rows()[0]["users"], 1)

    def test_manual_archive_timestamps_do_not_invent_load_dates(self):
        self.marker(last_use=iso(self.now), reported=iso(self.now), interval_minutes=5)
        self.scanner.scan_path(self.source, "counter")
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["day"]["users"], 0)
        self.assertEqual(self.store.monthly_rows()[0]["users"], 1)

    def test_configured_share_imports_explicit_daily_evidence(self):
        self.marker(loaded_dates=["2026-10-28", "2026-10-30"])
        settings = replace(self.settings, counter_paths=(self.source,))
        scanner = Scanner(settings, self.store)
        self.assertEqual(scanner.scan_all()[0]["records"], 1)
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["day"]["users"], 1)
        self.now += timedelta(hours=3)
        scanner.scan_all()
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["day"]["users"], 1)
        self.assertEqual(self.store.monthly_rows()[0]["observed_user_days"], 2)

    def test_manual_historical_import_restores_only_explicit_original_days(self):
        self.marker(month="2026-09", loaded_dates=["2026-09-15", "2026-09-17"])
        self.scanner.scan_path(self.source, "counter")
        self.assertEqual(self.store.dashboard("2026-09", "2026-09-15")["day"]["users"], 1)
        self.assertEqual(self.store.dashboard("2026-09", "2026-09-16")["day"]["users"], 0)
        self.assertEqual(self.store.dashboard("2026-10", "2026-10-30")["day"]["users"], 0)

    def test_invalid_json_does_not_block_other_valid_files(self):
        self.marker()
        (self.source / "broken.json").write_text("{unfinished")
        result = self.scanner.scan_path(self.source, "counter")
        self.assertEqual(result["records"], 1)
        self.assertEqual(result["errors"], 1)
        self.assertEqual(self.store.monthly_rows()[0]["users"], 1)
        self.assertIn("no filename fallback", result["detail"])

    def test_oversize_source_is_explicit_error(self):
        self.marker()
        scanner = Scanner(replace(self.settings, maximum_file_bytes=10), self.store)
        result = scanner.scan_path(self.source, "counter")
        self.assertEqual(result["records"], 0)
        self.assertEqual(result["errors"], 1)
        self.assertEqual(self.store.monthly_rows(), [])

    def test_log_append_and_unchanged_rescans_do_not_double_count(self):
        path = self.log(text=log_line() + log_line())
        self.scanner.scan_path(self.source, "log")
        self.assertEqual(self.store.log_summary()["records"], 2)
        self.scanner.scan_path(self.source, "log")
        self.assertEqual(self.store.log_summary()["records"], 2)
        with path.open("a", encoding="utf-8") as target:
            target.write(log_line(timestamp="2026-10-30 14:15:17.123"))
        self.scanner.scan_path(self.source, "log")
        self.assertEqual(self.store.log_summary()["records"], 3)

    def test_same_file_copied_to_another_root_is_not_double_counted(self):
        path = self.log()
        other = self.directory / "copied"
        other.mkdir()
        shutil.copy2(path, other / path.name)
        self.scanner.scan_path(self.source, "log")
        self.scanner.scan_path(other, "log")
        self.assertEqual(self.store.log_summary()["records"], 1)

    def test_copied_log_with_different_bom_and_newlines_is_not_double_counted(self):
        path = self.log(text=log_line() + log_line())
        other = self.directory / "text-converted-copy"
        other.mkdir()
        (other / path.name).write_bytes((log_line() + log_line()).replace("\n", "\r\n").encode("utf-8"))
        self.scanner.scan_path(self.source, "log")
        self.scanner.scan_path(other, "log")
        self.assertEqual(self.store.log_summary()["records"], 2)

    def test_valid_looking_incomplete_tail_waits_for_newline(self):
        path = self.log(text=log_line().rstrip("\n"))
        first = self.scanner.scan_path(self.source, "log")
        self.assertEqual(first["errors"], 1)
        self.assertEqual(self.store.log_summary()["records"], 0)
        with path.open("a", encoding="utf-8") as target:
            target.write(".\n")
        second = self.scanner.scan_path(self.source, "log")
        self.assertEqual(second["errors"], 0)
        self.assertEqual(self.store.log_summary()["records"], 1)

    def test_two_machine_logs_with_identical_invocations_are_distinct(self):
        self.log(machine="PC1")
        self.log(machine="PC2")
        self.scanner.scan_path(self.source, "log")
        self.assertEqual(self.store.log_summary()["records"], 2)
        self.assertEqual(self.store.log_summary()["users"], 1)

    def test_partial_log_tail_is_retried_without_duplicate_prefix(self):
        second = log_line(timestamp="2026-10-30 14:15:17.123")
        path = self.log(text=log_line() + second[:20])
        first = self.scanner.scan_path(self.source, "log")
        self.assertEqual(first["errors"], 1)
        self.assertEqual(self.store.log_summary()["records"], 1)
        with path.open("a", encoding="utf-8") as target:
            target.write(second[20:])
        second_result = self.scanner.scan_path(self.source, "log")
        self.assertEqual(second_result["errors"], 0)
        self.assertEqual(self.store.log_summary()["records"], 2)

    def test_recursion_is_explicit(self):
        nested = self.source / "child"
        nested.mkdir()
        (nested / "marker.json").write_text(json.dumps(usage()))
        self.assertEqual(self.scanner.scan_path(self.source, "counter", recursive=False)["records"], 0)
        self.assertEqual(self.scanner.scan_path(self.source, "counter", recursive=True)["records"], 1)

    def test_file_limit_is_reported(self):
        for index in range(3): self.marker(name=f"{index}.json", user=f"user-{index}")
        scanner = Scanner(replace(self.settings, maximum_files_per_scan=2), self.store)
        result = scanner.scan_path(self.source, "counter")
        self.assertEqual(result["records"], 2)
        self.assertEqual(result["errors"], 1)
        self.assertIn("File-count limit", result["detail"])


if __name__ == "__main__": unittest.main()
