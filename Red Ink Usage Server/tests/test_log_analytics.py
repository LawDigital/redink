"""Log feature analytics regressions; uses synthetic records, no user prompts."""
import json
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path

from fixtures import log_line
from redink_server.config import Settings
from redink_server.log_parser import parse_log_line
from redink_server.log_taxonomy import load_taxonomy, classify, TaxonomyError
from redink_server.scanner import Scanner
from redink_server.store import Store


class LogAnalyticsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.now = datetime(2026, 10, 31, 12, tzinfo=timezone.utc)
        self.settings = Settings(data_dir=self.root / "server", credentials={})
        self.store = Store(self.settings, clock=lambda: self.now)
        self.scanner = Scanner(self.settings, self.store)

    def tearDown(self):
        self.scanner.stop()
        self.temp.cleanup()

    def test_agent_classification_is_explicit_only(self):
        taxonomy = load_taxonomy(self.settings.data_dir)
        for label in ("RI_AutoPilot_Click", "AutoPilot (Mail)", "AutoPilot (Local)", "RI_WebAgent_Click", "AgentToolCall_Outlook_AutoPilot_m365_search"):
            self.assertTrue(classify(label,taxonomy)["agentic"], label)
        for label in ("LocalChat_SelectTools", "LocalChat_SetTooling", "LocalChat_Send", "RI_FreestyleAM_Click"):
            self.assertFalse(classify(label,taxonomy)["agentic"], label)

    def test_live_additions_preserve_counts_and_history(self):
        folder = self.root / "input"
        folder.mkdir()
        logfile = folder / "redink-0123456789ABCDEF-OL.log"
        logfile.write_text(log_line("RI_AutoPilot_Click invoked",timestamp="2026-10-30 12:00:00.000")
                          + log_line("AutoPilot (Mail) invoked",timestamp="2026-10-30 12:01:00.000")
                          + log_line("AgentRun_Outlook_AutoPilot invoked",timestamp="2026-10-30 12:01:30.000")
                          + log_line("LocalChat_SelectTools invoked",timestamp="2026-10-30 12:02:00.000")
                          + log_line("Diagnostic: secret prompt xyz",timestamp="2026-10-30 12:03:00.000"),encoding="utf-8")
        self.scanner.scan_path(folder,"log")
        data=self.store.log_analysis("30d","OL")
        self.assertEqual(data["invocations"],3)
        self.assertEqual(data["agentic_invocations"],1)
        self.assertEqual(data["standard_invocations"],2)
        self.assertEqual(data["other_entries"],1)
        self.assertNotIn("secret",str(data))
        self.assertIn("AutoPilot (Mail)",[r["feature"] for r in data["legacy_markers"]])
        self.assertIn("AgentRun_Outlook_AutoPilot",[r["feature"] for r in data["features"]])
        self.assertEqual(self.store.log_analysis("30d","WD")["invocations"],0)
        self.assertEqual(self.store.log_analysis("7d","ALL")["invocations"],3)
        self.assertEqual(self.store.log_analysis("30d","ALL")["days"],[
            {"date":"2026-10-30","count":3,"agentic":1,"standard":2}])
        self.scanner.scan_path(folder,"log")
        self.assertEqual(self.store.log_analysis("all")["invocations"],3)
        with logfile.open("a",encoding="utf-8") as f:
            f.write(log_line("RI_QuickTranslate_Click invoked",timestamp="2026-10-31 13:00:00.000"))
        self.scanner.scan_path(folder,"log")
        self.assertEqual(self.store.log_analysis("30d")["invocations"],4)
        logfile.unlink()
        self.scanner.scan_path(folder,"log")
        self.assertEqual(self.store.log_analysis("30d")["invocations"],4)
        self.assertEqual(self.store.dashboard("2026-10","2026-10-31")["month"]["users"],0)

    def test_chunked_growth_and_file_replacement(self):
        from dataclasses import replace
        folder=self.root / "big"
        folder.mkdir()
        path=folder / "redink-0123456789ABCDEF-WD.log"
        lines=[log_line("RI_PrimLang2_Click invoked",timestamp=f"2026-10-30 12:{i//60:02d}:{i%60:02d}.000")
               for i in range(950)]
        path.write_text("".join(lines),encoding="utf-8")
        scanner=Scanner(replace(self.settings,maximum_file_bytes=65536),self.store)
        first=scanner.scan_path(folder,"log")
        self.assertLess(first["records"],len(lines))
        self.assertIn("pending",first["detail"])
        for _ in range(4):scanner.scan_path(folder,"log")
        self.assertEqual(self.store.log_analysis("all")["invocations"],950)
        scanner.scan_path(folder,"log")
        self.assertEqual(self.store.log_analysis("all")["invocations"],950)
        path.write_text(log_line("RI_QuickTranslate_Click invoked",timestamp="2026-10-31 10:00:00.000"),encoding="utf-8")
        scanner.scan_path(folder,"log")
        self.assertEqual(self.store.log_analysis("all")["invocations"],951)

    def test_upgrade_of_previous_event_ids_does_not_double_count(self):
        import hashlib
        folder=self.root / "older"
        folder.mkdir()
        path=folder / "redink-0123456789ABCDEF-OL.log"
        line=log_line("AutoPilot (Mail) invoked",timestamp="2026-10-30 12:00:00.000")
        path.write_text(line,encoding="utf-8")
        event=parse_log_line(line,"OL")
        fingerprint=hashlib.sha256(line.strip().encode("utf-8")).hexdigest()
        old=hashlib.sha256((path.name.casefold()+"|OL|"+fingerprint+"|1").encode()).hexdigest()
        previous=dict(event, feature="AutoPilot")
        self.store.add_log_records([(old,previous)],str(path))
        self.store.save_file_status("log",path,path.stat(),1,0,0,"")
        self.assertEqual(self.store.log_analysis("all")["invocations"],1)
        self.scanner.scan_path(folder,"log")
        data=self.store.log_analysis("all")
        self.assertEqual(data["raw_invocations"],1)
        self.assertEqual(data["agentic_invocations"],0)
        self.assertEqual(data["legacy_markers"],[{"feature":"AutoPilot (Mail)","count":1}])
        with path.open("a",encoding="utf-8") as stream:
            stream.write(log_line("AutoPilot (Local) invoked",timestamp="2026-10-30 12:01:00.000"))
        self.scanner.scan_path(folder,"log")
        self.assertEqual(self.store.log_analysis("all")["raw_invocations"],2)

    def test_partial_logs_empty_range_and_validation(self):
        self.assertEqual(self.store.log_analysis("30d")["entries"],0)
        self.assertEqual(self.store.log_analysis("30d")["days"],[])
        for period, host in (("1d","ALL"),("all","XX"),("' DROP TABLE","OL")):
            with self.assertRaises(ValueError): self.store.log_analysis(period,host)

    def test_taxonomy_customization_and_invalid_file(self):
        root = self.settings.data_dir
        root.mkdir(exist_ok=True)
        override = root / "feature_taxonomy.json"
        override.write_text(json.dumps({"groups":[{"id":"workflow","label":"Workflow","agentic":True,"patterns":["MyTask_*"]}]}))
        self.assertTrue(classify("MyTask_Run",load_taxonomy(root))["agentic"])
        self.assertFalse(classify("RI_AutoPilot_Click",load_taxonomy(root))["agentic"])
        override.write_text('{bad',encoding='utf-8')
        with self.assertRaises(TaxonomyError):load_taxonomy(root)


if __name__ == "__main__": unittest.main()
