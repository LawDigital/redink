"""Cross-language contract using actual production VB builder output."""
import json
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import parse_qs

from redink_server.config import Settings
from redink_server.protocol import normalize_event, parse_legacy_path, parse_legacy_query
from redink_server.store import Store


class VbVectorTests(unittest.TestCase):
    def test_actual_vb_json_query_form_and_file_preserve_sparse_dates_and_monthly_identity(self):
        vectors = json.loads((Path(__file__).parent / "data" / "counter-vb-vectors-v2.json").read_text())
        json_event = normalize_event(vectors["json"])
        events = [json_event, normalize_event(vectors["full_record"]), parse_legacy_query(vectors["query"]), normalize_event(parse_qs(vectors["form"]))]
        for event in events:
            self.assertEqual(event["loaded_dates"], ["2026-10-01", "2026-10-29", "2026-10-30"])
            self.assertEqual(event["loaded_hosts"], ["EX", "OL", "WD"])
            self.assertEqual(event["billing_event_id"], json_event["billing_event_id"])
            self.assertEqual(event["host_event_id"], json_event["host_event_id"])
        legacy = parse_legacy_path(vectors["path"])
        self.assertEqual(legacy["billing_event_id"], json_event["billing_event_id"])
        self.assertEqual(legacy["loaded_dates"], [])
        with tempfile.TemporaryDirectory() as directory:
            store = Store(Settings(Path(directory), {}), clock=lambda: datetime(2026, 10, 30, 13, tzinfo=timezone.utc))
            for event in [legacy] + events + events:
                store.ingest(event)
            row = store.monthly_rows()[0]
            self.assertEqual(row["users"], 1)
            self.assertEqual(row["observed_user_days"], 3)
            self.assertEqual((row["word_users"], row["excel_users"], row["outlook_users"]), (1, 1, 1))
            self.assertEqual(store.dashboard("2026-10", "2026-10-02")["day"]["users"], 0)
            self.assertEqual(store.dashboard("2026-10", "2026-10-30")["day"]["users"], 1)


if __name__ == "__main__": unittest.main()
