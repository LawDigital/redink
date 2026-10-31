import tempfile
import unittest
from pathlib import Path
from redink_server.log_parser import parse_log_line, iter_log_records, log_host_from_name


def log_line(message="Correct_Outlook invoked", version="V2.3.1", user="0123456789ABCDEF", timestamp="2026-10-30 14:15:16.123"):
    return f"[{timestamp}] {user} {version:<20} {message}\n"


class LogTests(unittest.TestCase):
    def test_exact_writer_format(self):
        row = parse_log_line(log_line(), "OL")
        self.assertEqual(row["date"], "2026-10-30")
        self.assertEqual(row["host"], "OL")
        self.assertEqual(row["user_hash"], "0123456789ABCDEF")
        self.assertEqual(row["version"], "V2.3.1")
        self.assertEqual(row["feature"], "Correct_Outlook")
        self.assertTrue(row["invoked"])
        self.assertFalse(row["timestamp_local"].endswith("Z"))

    def test_agent_labels_preserved_without_raw_messages(self):
        self.assertEqual(parse_log_line(log_line("AutoPilot (Mail) invoked"))["feature"], "AutoPilot (Mail)")
        self.assertEqual(parse_log_line(log_line("AgentToolCall_Outlook_AutoPilot_m365_search invoked."))["feature"], "AgentToolCall_Outlook_AutoPilot_m365_search")

    def test_empty_and_long_versions(self):
        row = parse_log_line(log_line(version=""))
        self.assertEqual(row["version"], "")
        self.assertTrue(row["invoked"])
        version = "V2026.10.30.12345678901234567890"
        self.assertEqual(parse_log_line(log_line(version=version))["version"], version)

    def test_diagnostic_content_not_retained(self):
        row = parse_log_line(log_line("Unknown voice command: SECRET CUSTOMER NAMES"))
        self.assertFalse(row["invoked"])
        self.assertEqual(row["feature"], "Other activity")
        self.assertNotIn("SECRET", str(row))
        self.assertNotIn("CUSTOMER", str(row))

    def test_malformed_lines(self):
        self.assertIsNone(parse_log_line("Error: prompt=PRIVATE"))
        self.assertIsNone(parse_log_line(log_line(timestamp="2026-13-30 14:15:16.123")))
        self.assertIsNone(parse_log_line(log_line(user="alice")))

    def test_filename_filter_excel_mapping(self):
        self.assertEqual(log_host_from_name("redink-1234567890ABCDEF-XL.log"), "EX")
        self.assertIsNone(log_host_from_name("redink-logger-debug.log"))
        self.assertIsNone(log_host_from_name("RedInk_Agent_Conversation.txt"))

    def test_iterator_bom_duplicates_and_errors(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "redink-1234567890ABCDEF-XL.log"
            path.write_text(log_line() + log_line() + "bad\n", encoding="utf-8-sig")
            rows = list(iter_log_records(path))
            self.assertEqual(len(rows), 3)
            self.assertEqual(rows[0].event["host"], "EX")
            self.assertEqual(rows[1].line, 2)
            self.assertIsNotNone(rows[2].error)
            self.assertIsNone(rows[2].event)
            self.assertEqual(list(iter_log_records(path, max_file_bytes=5))[0].line, 0)


if __name__ == "__main__": unittest.main()
