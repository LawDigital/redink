"""Regression fixtures for primary accounting, diagnostic framing and lossless log archival."""
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from datetime import date, datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
from zipfile import ZipFile

from fixtures import log_line
from redink_server.archive_logs import archive_segments, month_boundary
from redink_server.browser_start import browser_requested, dashboard_url
from redink_server.config import Settings
from redink_server.log_parser import iter_log_records, log_host_from_name, parse_log_line
from redink_server.scanner import Scanner
from redink_server.store import Store
from redink_server.usage_semantics import describe
from redink_server.user_analytics import function_users, user_summary_csv

ROOT = Path(__file__).resolve().parents[1]
TODAY = date(2026,10,31)

class UsageV14Tests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(); self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name); self.logs = self.root/'incoming'; self.logs.mkdir()
        self.setting = Settings(data_dir=self.root/'database',credentials={},log_paths=(self.logs,),maximum_file_bytes=65536)
        self.store = Store(self.setting,clock=lambda:datetime(2026,10,31,12,tzinfo=timezone.utc))
        self.scanner = Scanner(self.setting,self.store); self.addCleanup(self.scanner.stop)
        self.path = self.logs/'redink-0123456789ABCDEF-202609-WD.log'
    def scan(self, **kwargs): return self.scanner.scan_path(self.logs,'log',**kwargs)
    def archive(self, apply=False): return archive_segments(self.store,'2026-10',self.root/'archive',apply=apply,today=TODAY)
    def test_exclusive_classes_and_compatible_aliases(self):
        for name in ('tool_loader','tool_describe','context_expand','report_progress','skill_use'):
            self.assertEqual(describe('AgentToolCall_Word_Agent_'+name).kind,'internal')
        for k,name in [('agent','agent_workspace_file_reader'),('tool','agent_workspace_read'),('skill','skill_résumé')]:
            key='AgentToolCall_Word_Agent_'+name
            self.assertEqual(describe(parse_log_line(log_line(key+' invoked'))['feature']).kind,k)
        for old in ('RI_BallooMergePart_Click','BallooMergePart_Word','BalloonMergePart_Word'):
            self.assertEqual(describe(old).label,'BalloonMergePart_Word')
        self.assertTrue(describe('AgentRun_Outlook_Tooling').primary)
        self.assertFalse(describe('AgentToolCall_unknown_tool').primary)
        self.assertFalse(describe('AutoPilot (Mail)').primary)
    def test_store_name_is_label_not_address_or_old_id(self):
        key='AgentToolCall_Outlook_LocalChatAgent_knowledge_search_store'
        raw=parse_log_line(log_line(key+' (Zürich – Steuern) invoked'))
        self.assertEqual(describe(raw['feature']).label,'Knowledge search · Zürich – Steuern')
        self.assertNotIn('ab123',describe(key+'_ab123').label)
        self.assertIsNone(parse_log_line(log_line(key+' (bad\nname) invoked')))
    def test_user_rates_period_partial_weeks_distinct_hosts_and_median(self):
        events=[('2026-10-25','Chat_Word'),('2026-10-26','AgentRun_Word_Agent'),('2026-10-26','Chat_Word'),('2026-10-26','AgentToolCall_Word_Agent_tool_loader'),('2026-10-26','AgentToolCall_Word_Agent_excel_read')]
        self.path.write_text(''.join(log_line(f+' invoked',timestamp=d+' 12:00:00.000') for d,f in events),'utf-8');self.scan()
        data=function_users(self.store,'7d','ALL'); user=data['users'][0]
        self.assertEqual((data['period_days'],user['invocations'],user['calls_per_active_day'],user['calls_per_calendar_day'],user['calls_per_7_days'],user['calls_per_active_week'],user['active_weeks'],user['peak_day_count'],user['median_per_active_day']),(7,3,1.5,.43,3,1.5,2,2,1.5))
        self.assertIn('Per active ISO week',user_summary_csv(data))
        d=function_users(self.store,'7d','ALL',user['user_hash'])['selected'];self.assertEqual(d['tool_details'],[{'feature':'excel_read','count':1}]);self.assertEqual(sum(w['count'] for w in d['weeks']),3)
        self.assertEqual(function_users(self.store,'all')['period_days'],2)
        self.assertEqual(function_users(self.store,'30d','OL')['total_users'],0)
        analytics=self.store.log_analysis('all');self.assertEqual(analytics['raw_invocations'],5); self.assertEqual(analytics['agentic_invocations'],1)
        self.assertEqual(sum(r['count'] for g in analytics['families'] for r in g['commands']),3)
        self.assertEqual(self.store.dashboard('2026-10','2026-10-26')['month']['users'],0)
    def test_copy_and_status_json_filtered_before_budget(self):
        self.path.write_text(log_line(),'utf-8');shutil.copyfile(self.path,self.path.with_name(self.path.stem+'(1).log'))
        (self.logs/'redink-autopilot-status-example.json').write_text('invalid','utf-8');result=self.scan()
        self.assertEqual((result['files'],result['errors'],result['detail']),(2,0,''));self.assertEqual(self.store.log_analysis('all')['entries'],1)
        self.assertEqual(log_host_from_name('redink-0123456789ABCDEF-202610-XL.log'),'EX')
    def test_chunk_resume_diagnostic_and_explicit_bad_record(self):
        # Move the first continuation to the next scan, past the committed 64 KiB budget.
        line=log_line(); block=line*((65536-100)//len(line)); header=log_line('Surgical source recovery requested: '+'X'*240)
        self.path.write_text(block+header+'   bei Synthetic.Op() in PRIVATE PATH\n--- End of inner exception stack trace ---\n'+line+'   at MalformedAfterInvocation()\n','utf-8')
        one=self.scan();self.assertIn('pending',one['detail']);two=self.scan();self.assertEqual(two['errors'],1);self.assertNotIn('PRIVATE',two['detail']);self.assertEqual(two['skipped'],2)
        self.assertEqual(sum(bool(x.error) for x in iter_log_records(self.path)),1)
        self.assertNotIn('MalformedAfterInvocation',str(self.store.log_analysis('all')))
    def test_legacy_cursor_replay_does_not_duplicate(self):
        self.path.write_text(log_line()+log_line('diagnostic')+'   at Function()\n','utf-8');self.scan()
        with self.store.connect() as con:
            tables=[r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")];table=next(t for t in tables if t.startswith('log_') and 'cursor' in t)
            con.execute('UPDATE '+table+" SET head_signature='old-parser'")
        self.assertEqual(self.scan()['errors'],0);self.assertEqual(self.store.log_analysis('all')['entries'],2)
    def test_archive_preview_verified_bytes_apply_then_replay(self):
        blob=log_line().replace('\n','\r\n').encode('utf-8-sig'); self.path.write_bytes(blob);self.scan()
        self.assertEqual(self.archive()['eligible'],1);self.assertTrue(self.path.exists());self.assertFalse((self.root/'archive').exists())
        self.assertEqual(self.archive(True)['archived'],1);self.assertFalse(self.path.exists())
        with ZipFile(next((self.root/'archive').glob('*.zip'))) as z:self.assertEqual(z.read(self.path.name),blob)
        self.path.write_bytes(blob);self.scan();self.assertEqual(self.store.log_analysis('all')['entries'],1)
    def test_archive_never_discards_unseen_or_pruned_data(self):
        self.path.write_text(log_line(),'utf-8');self.assertEqual(self.archive(True)['held'],1)
        self.scan()
        with self.store.connect() as con:con.execute('DELETE FROM log_events')
        self.assertEqual(self.archive(True)['held'],1); self.assertTrue(self.path.exists())
    def test_archive_excludes_current_unsplit_and_unsafe_destination(self):
        self.assertEqual(month_boundary(6,TODAY),'2026-05'); self.assertEqual(month_boundary(2,date(2026,1,1)),'2025-12')
        for n in ['redink-0123456789ABCDEF-WD.log','redink-0123456789ABCDEF-202610-WD.log']:(self.logs/n).write_text(log_line(),'utf-8')
        self.scan();self.assertFalse(self.archive(True)['files']);self.assertEqual(len(list(self.logs.iterdir())),2)
        with self.assertRaises(ValueError): archive_segments(self.store,'2026-10',self.logs/'archives',today=TODAY)
    def test_safe_browser_url_and_headless_override(self):
        self.assertEqual(dashboard_url(self.setting),'http://127.0.0.1:8765/')
        for host,trusted,url in [('::',[],'https://localhost:443/'),('::1',[],'https://[::1]:443/'),('0.0.0.0',['.example','srv.example'],'https://srv.example:443/')]:
            self.assertEqual(dashboard_url(SimpleNamespace(host=host,tls_enabled=True,port=443,trusted_hosts=trusted)),url)
        with self.assertRaises(ValueError): dashboard_url(SimpleNamespace(host='evil/path',tls_enabled=False,port=8,trusted_hosts=[]))
        stream=Mock();stream.isatty.return_value=False;self.assertFalse(browser_requested(None,stream));self.assertTrue(browser_requested(True,stream));stream.isatty.return_value=True;self.assertTrue(browser_requested(None,stream));self.assertFalse(browser_requested(False,stream))
    def test_build_helper_in_path_with_spaces(self):
        project=self.root/'Red Ink Usage Server';project.mkdir();helper=project/'build support.py';shutil.copyfile(ROOT/'scripts/build_support.py',helper)
        def call(*args):return subprocess.run([sys.executable,str(helper),*map(str,args)],check=True,capture_output=True,text=True).stdout
        self.assertIn('Architecture:',call('check-python'));self.assertIn(str(sys.version_info.major)+'.',call('python-version'))
        target=self.root/'release output.zip'; call('archive',project,target)
        with ZipFile(target) as z:self.assertEqual(z.read(project.name+'/'+helper.name),helper.read_bytes())
        build=(ROOT/'build-windows.ps1').read_text('utf-8-sig');self.assertNotIn("@('-c',",build)
