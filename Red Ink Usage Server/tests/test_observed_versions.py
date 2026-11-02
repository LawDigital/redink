"""Versions are distinct log users, not event counts, with one per host."""
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path

from redink_server.config import Settings
from redink_server.store import Store
from redink_server.version_comparison import version_sort_key


class VersionAggregationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.store = Store(Settings(data_dir=Path(self.tmp.name), credentials={}),
                           clock=lambda: datetime(2026, 11, 2, 12, tzinfo=timezone.utc))
        self.records = []

    def add(self, user, host, version, day, suffix='1', invoked=1):
        # Each event is uniquely identified; repeated events must not inflate users.
        stamp = f'{day}T12:00:00.{int(suffix):03d}'
        self.records.append((f'{user}:{host}:{version}:{day}:{suffix}', dict(
            date=day, timestamp_local=stamp, host_id=host, user_hash=user,
            version=version, feature='Translate_Word', invoked=invoked)))

    def commit(self):
        self.store.add_log_records(self.records, 'synthetic')

    def test_latest_semantic_per_user_and_host_and_no_event_double_count(self):
        for i in range(12):
            self.add('A', 'WD', 'V2.3.1', '2026-11-01', str(i))
        self.add('A', 'WD', 'V2.4.0', '2026-11-01', '80')
        # Version 2.3.1 appears again AFTER the upgrade, but 2.4.0 is newer.
        self.add('A', 'WD', 'V2.3.1', '2026-11-02', '81')
        self.add('B', 'WD', 'V2.3.1', '2026-11-02')
        self.add('A', 'EX', 'V2.3.1', '2026-11-02')
        self.add('A', 'OL', 'V2.4.0', '2026-11-02')
        self.add('A', 'OL', '', '2026-11-02', '12')
        self.add('C', 'WD', '', '2026-11-02')
        self.add('A', 'UNKNOWN', 'V2.4.0', '2026-11-02')
        self.commit()
        got = self.store.recent_addin_versions()
        self.assertEqual({(r['host'],r['version']):r['count'] for r in got['versions']},
                         {('WD','V2.4.0'):1, ('WD','V2.3.1'):1,
                          ('EX','V2.3.1'):1, ('OL','V2.4.0'):1, ('WD','Version unavailable'):1})
        self.assertEqual(sum(r['count'] for r in got['versions']), 5)
        self.assertEqual(self.store.log_analysis('all')['version_summary']['versions'], got['versions'])
        self.assertEqual(self.store.recent_addin_versions('7d','WD')['users_by_addin'],
                         [{'host':'WD','count':3}])

    def test_short_window_independent_of_general_log_analysis(self):
        self.add('A','WD','V2.2.0','2026-10-05')
        self.add('B','WD','V2.3.0','2026-10-24')
        self.add('C','WD','V2.4.0','2026-11-02')
        self.commit()
        self.assertEqual(sum(v['count'] for v in self.store.log_analysis('all')['versions']),1)
        self.assertEqual(sum(v['count'] for v in self.store.log_analysis('30d','ALL','14d')['versions']),2)
        self.assertEqual(sum(v['count'] for v in self.store.log_analysis('7d','ALL','30d')['versions']),3)
        for period in ('all','365d','wrong'):
            with self.assertRaises(ValueError): self.store.recent_addin_versions(period)

    def test_date_style_and_unknown_format_fallback(self):
        self.assertLess(version_sort_key('V.311025'), version_sort_key('V.011125'))
        self.add('A','OL','V.311025','2026-11-02')
        self.add('A','OL','V.011125','2026-11-01')
        # Unrecognised schemes: use latest log observation, not arbitrary lexicographic sort.
        self.add('B','WD','BuildAlpha','2026-11-01')
        self.add('B','WD','BuildBeta','2026-11-02')
        self.commit()
        got = self.store.recent_addin_versions()
        self.assertEqual({(r['host'],r['version']) for r in got['versions']},
                         {('OL','V.011125'), ('WD','BuildBeta')})
