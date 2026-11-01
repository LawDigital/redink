import json
import tempfile
import unittest
from pathlib import Path
from redink_server.importers import iter_counter_records, safe_iter_files
from test_protocol import usage, compact, path_event


class ImportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
    def tearDown(self): self.temp.cleanup()
    def write(self, name, text):
        path = self.folder / name
        path.write_text(text, encoding="utf-8-sig")
        return path

    def test_marker_and_jsonl(self):
        path = self.write("marker.json", json.dumps(usage(), indent=2))
        self.assertEqual(list(iter_counter_records(path))[0].event["month"], "2026-09")
        path = self.write("combined.log", json.dumps(usage()) + "\n" + json.dumps(usage(hostId="EX")) + "\n")
        self.assertEqual([r.event["host_id"] for r in iter_counter_records(path)], ["WD", "EX"])

    def test_json_array_and_access_logs(self):
        path = self.write("events.json", json.dumps([usage(), usage()]))
        self.assertEqual(len(list(iter_counter_records(path))), 2)
        path = self.write("access.log", '#Fields: date time cs-method cs-uri-stem\n127.0.0.1 - - "GET ' + path_event() + ' HTTP/1.1" 200 2\n')
        rows = list(iter_counter_records(path))
        self.assertTrue(rows[0].skipped)
        self.assertEqual(rows[1].event["month"], "2026-09")

    def test_corrupt_marker_no_filename_fallback(self):
        path = self.write("RI_LC_2026-09_n0123456789abcdef_PRO_WD_u" + "1"*32 + ".json", "{partial")
        row = list(iter_counter_records(path))[0]
        self.assertIsNone(row.event)
        self.assertIn("no filename fallback", row.error)

    def test_state_and_unrelated_file_not_usage(self):
        path = self.write("RI_LC_State.json", '{"v": 1, "knownHostEventIds":{}}')
        self.assertTrue(list(iter_counter_records(path))[0].skipped)
        path = self.write("unrelated.log", "nothing relevant here\n")
        self.assertTrue(list(iter_counter_records(path))[0].skipped)

    def test_periodic_marker_timestamps_survive_import(self):
        event = usage(lastUseUtc="2026-10-30T10:11:12Z", reportedUtc="2026-10-30T10:12:12Z")
        path = self.write("periodic-marker.json", json.dumps(event))
        row = list(iter_counter_records(path))[0]
        self.assertIsNone(row.error)
        self.assertEqual(row.event["last_use_utc"], "2026-10-30T10:11:12.000000Z")
        self.assertEqual(row.event["reported_utc"], "2026-10-30T10:12:12.000000Z")

    def test_duplicate_fields_and_limits_explicit(self):
        path = self.write("duplicate.json", '{"t":"U","t":"E"}')
        self.assertEqual(list(iter_counter_records(path))[0].error, "Duplicate JSON field.")
        path = self.write("large.json", json.dumps(usage()))
        self.assertIn("exceeds", list(iter_counter_records(path, max_file_bytes=2))[0].error)
        path = self.write("long.log", "x"*100 + "\n" + json.dumps(usage()) + "\n")
        rows = list(iter_counter_records(path, max_line_bytes=40))
        self.assertEqual(len(rows), 2)
        self.assertTrue(all(row.error for row in rows))

    def test_symlinks_do_not_escape(self):
        nested = self.folder / "inside"; nested.mkdir()
        self.write("good.json", json.dumps(usage()))
        outside = self.folder.parent / (self.folder.name + "-outside"); outside.mkdir()
        try:
            (outside / "other.json").write_text(json.dumps(usage()))
            try:
                (nested / "escape").symlink_to(outside, target_is_directory=True)
                (nested / "escape.json").symlink_to(outside / "other.json")
            except OSError: self.skipTest("Symlink privileges unavailable")
            files = list(safe_iter_files(self.folder))
            self.assertEqual(len(files), 1)
            self.assertEqual(files[0].name, "good.json")
            self.assertIn("Symbolic", list(iter_counter_records(nested / "escape.json"))[0].error)
        finally:
            (outside / "other.json").unlink(); outside.rmdir()

    def test_deeply_nested_json_is_explicit_parse_error(self):
        path = self.write("nested.json", "[" * 2000 + "0" + "]" * 2000)
        row = list(iter_counter_records(path))[0]
        self.assertIsNone(row.event)
        self.assertIsNotNone(row.error)

    def test_missing_root_is_outage(self):
        with self.assertRaises(OSError): list(safe_iter_files(self.folder / "missing"))


if __name__ == "__main__": unittest.main()
