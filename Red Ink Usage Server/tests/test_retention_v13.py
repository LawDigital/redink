"""Retention safety: cleanup, import replay protection, cutoff boundaries, and snapshots."""
from __future__ import annotations
import json
import sqlite3
import tempfile
import unittest
from dataclasses import replace
from datetime import datetime, timezone
from pathlib import Path

from fixtures import usage, log_line, log_filename
from redink_server.config import Settings, initialize, load_settings, ConfigError
from redink_server.settings_ui import configuration_document, update_configuration
from redink_server.protocol import normalize_event
from redink_server.store import Store
from redink_server.retention import effective_cutoff
from redink_server.log_parser import parse_log_line

class RetentionTests(unittest.TestCase):
    def setUp(self):
        tmp=tempfile.TemporaryDirectory();self.addCleanup(tmp.cleanup)
        self.root=Path(tmp.name);initialize(self.root)
        self.settings=load_settings(self.root)
        self.now=datetime(2026,10,31,12,tzinfo=timezone.utc)
        self.store=Store(self.settings,clock=lambda:self.now)

    def event(self,month='2026-10',dates=None,host='WD',user='alice'):
        return normalize_event(usage(month=month,loaded_dates=dates,host=host,user=user))

    def cutoff(self, date='2026-10-15', months=0):
        self.settings=replace(self.settings,cutoff_date=date,keep_months=months)
        self.store=Store(self.settings,clock=lambda:self.now)

    def test_no_policy_keeps_existing_behavior(self):
        self.store.ingest(self.event('2026-07'))
        self.assertEqual(len(self.store.monthly_rows()),1)
        self.assertFalse(self.store.retention_preview()['enabled'])

    def test_before_cutoff_is_acknowledged_but_never_stored(self):
        self.cutoff()
        out=self.store.ingest(self.event('2026-09', ['2026-09-01']))
        self.assertTrue(out['accepted']);self.assertTrue(out['discarded'])
        self.assertEqual(self.store.monthly_rows(),[])
        self.assertEqual(self.store.ingest(self.event('2026-10',['2026-10-02']))['discarded'],True)
        self.assertEqual(self.store.monthly_rows(),[])

    def test_mixed_dates_keeps_one_monthly_user_and_only_new_days(self):
        self.cutoff()
        record=self.event('2026-10',['2026-10-02','2026-10-15','2026-10-30'])
        result=self.store.ingest(record)
        self.assertTrue(result['accepted']);self.assertEqual(result['loaded_dates_added'],2)
        rows=self.store.user_rows('2026-10')
        self.assertEqual(rows[0]['loaded_dates'],'2026-10-15,2026-10-30')
        self.assertEqual(self.store.monthly_rows()[0]['users'],1)

    def test_expired_legacy_boundary_and_new_legacy(self):
        self.cutoff()
        # First use before the boundary: legacy monthly marker cannot prove a newer load.
        self.assertTrue(self.store.ingest(self.event())['discarded'])
        newer=self.event();newer['first_use_utc']='2026-10-16T08:00:00Z'
        self.assertTrue(self.store.ingest(newer)['accepted'])
        self.assertEqual(self.store.monthly_rows()[0]['users'],1)

    def test_cleaning_removes_old_identity_hosts_logs_and_snapshots(self):
        self.store.ingest(self.event('2026-09',['2026-09-12']))
        self.store.ingest(self.event('2026-10',['2026-10-01','2026-10-20']))
        self.store.ingest(self.event('2026-10',['2026-10-02'],user='bob'))
        self.store.ingest(self.event('2026-10',['2026-10-20'],user='carol'))
        self.store.add_log_records([('old',{'date':'2026-10-01','timestamp_local':'2026-10-01 12:00:00.000','host_id':'WD','user_hash':'TEST','version':'v1','feature':'Correct','invoked':True}),
                                    ('new',{'date':'2026-10-20','timestamp_local':'2026-10-20 12:00:00.000','host_id':'WD','user_hash':'TEST','version':'v1','feature':'Correct','invoked':True})], 'fixture')
        self.store.save_report_snapshot({'to_month':'2026-09','from_month':'2026-09'})
        self.store.save_report_snapshot({'to_month':'2026-10','from_month':'2026-09'})
        self.cutoff()
        preview=self.store.retention_preview()
        self.assertEqual(preview['counts']['billing_usage'],2)
        self.assertEqual(preview['counts']['log_events'],1)
        cleaned=self.store.clean_retention('2026-10-15')
        self.assertTrue(cleaned['cleaned'])
        self.assertEqual(self.store.retention_preview()['total'],0)
        self.assertEqual(len(self.store.user_rows('2026-10')),2)
        self.assertEqual(self.store.daily_rows('2026-10-01','2026-10-14'),[])
        self.assertEqual(self.store.log_summary()['records'],1)
        with self.store.connect() as con:
            self.assertEqual(con.execute('SELECT COUNT(*) FROM report_snapshots').fetchone()[0],1)
            self.assertEqual(con.execute('PRAGMA foreign_key_check').fetchall(),[])
        self.assertTrue(self.store.ingest(self.event('2026-09'))['discarded'])
        self.assertEqual(self.store.clean_retention('2026-10-15')['remaining']['billing_usage'],0)

    def test_rolling_boundary_validates_calendar(self):
        from datetime import date
        self.assertEqual(effective_cutoff('2026-01-01',1,date(2026,3,31)),'2026-02-28')
        self.assertEqual(effective_cutoff('2026-05-01',12,date(2026,3,31)),'2026-05-01')
        self.assertEqual(effective_cutoff('',0,date(2026,3,31)),'')

    def test_config_includes_retention_and_rejects_invalid(self):
        doc=configuration_document(self.settings)
        self.assertEqual(doc['sections']['retention']['keep_months'],0)
        update_configuration(self.settings,{'retention':{'cutoff_date':'2026-10-15','keep_months':24,'auto_cleanup':True}})
        loaded=load_settings(self.root)
        self.assertEqual(loaded.cutoff_date,'2026-10-15');self.assertEqual(loaded.keep_months,24)
        with self.assertRaises((ConfigError,ValueError)):
            update_configuration(loaded,{'retention':{'cutoff_date':'2026-02-31'}})
        self.assertEqual(load_settings(self.root).cutoff_date,'2026-10-15')

    def test_imports_do_not_recreate_old_logs(self):
        self.cutoff()
        original={'date':'2026-09-01','timestamp_local':'2026-09-01 12:00:00.000','host_id':'OL','user_hash':'TEST','version':'v1','feature':'AutoPilot','invoked':True}
        self.assertEqual(self.store.add_log_records([('old',original)],'source'),0)
        self.assertEqual(self.store.log_summary()['records'],0)
        new=dict(original,date='2026-10-30')
        self.assertEqual(self.store.add_log_records([('new',new)],'source'),1)

if __name__=='__main__':unittest.main()
