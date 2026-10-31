"""Automatic upgrades from existing deployments, no manual migration command."""
from __future__ import annotations

import sqlite3
from contextlib import closing
import tempfile
from sqlite_tempdir import SqliteTemporaryDirectory
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

from redink_server.config import Settings
from redink_server.store import Store, CURRENT_SCHEMA_VERSION


class AutomaticMigrationTests(unittest.TestCase):
    def setUp(self):
        temp = SqliteTemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.settings = Settings(data_dir=self.root, credentials={})
        self.clock = lambda: datetime(2026, 10, 31, 10, tzinfo=timezone.utc)

    def test_new_install_creates_current_schema_without_unnecessary_backup(self):
        store = Store(self.settings, clock=self.clock)
        with store.connect() as con:
            self.assertEqual(con.execute('PRAGMA user_version').fetchone()[0], CURRENT_SCHEMA_VERSION)
            self.assertEqual(con.execute('SELECT version FROM schema_migrations').fetchone()[0], CURRENT_SCHEMA_VERSION)
        self.assertFalse(list((self.root/'backups').glob('*.sqlite3')) if (self.root/'backups').exists() else False)

    def test_v2_install_auto_migrates_preserves_data_and_creates_verified_backup(self):
        db = self.root / 'usage.sqlite3'
        with closing(sqlite3.connect(db)) as con, con:
            con.execute('CREATE TABLE server_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)')
            con.execute("INSERT INTO server_meta VALUES ('last_ingest_utc','2026-09-15T00:00:00Z')")
            con.execute('CREATE TABLE report_snapshots (report_id TEXT PRIMARY KEY, created_utc TEXT NOT NULL, content_json TEXT NOT NULL, sha256 TEXT NOT NULL)')
            con.execute("INSERT INTO report_snapshots VALUES ('RI-old','2026-08-01T00:00:00Z','{\"to_month\":\"2026-07\"}','abc')")
            con.execute('PRAGMA user_version=2')
        store = Store(self.settings, clock=self.clock)  # no explicit upgrade step
        with store.connect() as con:
            self.assertEqual(con.execute('PRAGMA user_version').fetchone()[0], 3)
            self.assertEqual(con.execute('SELECT value FROM server_meta WHERE key=\'last_ingest_utc\'').fetchone()[0], '2026-09-15T00:00:00Z')
            self.assertEqual(con.execute('SELECT report_id FROM report_snapshots').fetchone()[0], 'RI-old')
            self.assertEqual(con.execute('PRAGMA foreign_key_check').fetchall(), [])
        backups = list((self.root/'backups').glob('usage-before-schema-v2-to-v3-*.sqlite3'))
        self.assertEqual(len(backups), 1)
        with closing(sqlite3.connect(backups[0])) as con, con:
            self.assertEqual(con.execute('PRAGMA user_version').fetchone()[0], 2)
            self.assertEqual(con.execute('SELECT report_id FROM report_snapshots').fetchone()[0], 'RI-old')
        # Restarting a current version must not rewrite or create backups.
        Store(self.settings, clock=self.clock)
        self.assertEqual(len(list((self.root/'backups').glob('*.sqlite3'))), 1)

    def test_migration_backup_includes_committed_wal_transactions(self):
        db = self.root / 'usage.sqlite3'
        first = sqlite3.connect(db)
        try:
            first.execute('PRAGMA journal_mode=WAL')
            first.execute('CREATE TABLE server_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)')
            first.execute("INSERT INTO server_meta VALUES ('wal','present')")
            first.execute('PRAGMA user_version=2')
            first.commit()
            # Keep the original SQLite connection open while upgrading; the
            # online backup must include data committed to the WAL file.
            upgraded = Store(self.settings, clock=self.clock)
            with upgraded.connect() as con:
                self.assertEqual(con.execute("SELECT value FROM server_meta WHERE key='wal'").fetchone()[0], 'present')
            backups = list((self.root/'backups').glob('*.sqlite3'))
            self.assertEqual(len(backups), 1)
            with closing(sqlite3.connect(backups[0])) as con, con:
                self.assertEqual(con.execute("SELECT value FROM server_meta WHERE key='wal'").fetchone()[0], 'present')
        finally:
            first.close()

    def test_unversioned_existing_database_gets_a_backup_and_upgrade(self):
        with closing(sqlite3.connect(self.root/'usage.sqlite3')) as con, con:
            con.execute('CREATE TABLE server_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)')
            con.execute("INSERT INTO server_meta VALUES ('k','v')")
        store = Store(self.settings, clock=self.clock)
        with store.connect() as con:
            self.assertEqual(con.execute('SELECT value FROM server_meta WHERE key=\'k\'').fetchone()[0], 'v')
        self.assertEqual(len(list((self.root/'backups').glob('usage-before-schema-v0-to-v3-*.sqlite3'))), 1)

    def test_future_schema_never_downgrades(self):
        db = self.root/'usage.sqlite3'
        with closing(sqlite3.connect(db)) as con, con:
            con.execute('CREATE TABLE preserved (x TEXT)')
            con.execute("INSERT INTO preserved VALUES ('do not change')")
            con.execute('PRAGMA user_version=50')
        with self.assertRaisesRegex(RuntimeError, 'newer'):
            Store(self.settings, clock=self.clock)
        with closing(sqlite3.connect(db)) as con, con:
            self.assertEqual(con.execute('PRAGMA user_version').fetchone()[0], 50)
            self.assertEqual(con.execute('SELECT x FROM preserved').fetchone()[0], 'do not change')
        self.assertFalse((self.root/'backups').exists())

    def test_failed_upgrade_rolls_back_leaves_verified_backup(self):
        db = self.root/'usage.sqlite3'
        with closing(sqlite3.connect(db)) as con, con:
            con.execute('CREATE TABLE server_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)')
            con.execute("INSERT INTO server_meta VALUES ('existing','safe')")
            con.execute('PRAGMA user_version=2')
        # Force a constraint error inside the schema transaction with a
        # conflicting object. Must not partially create the new schema.
        with closing(sqlite3.connect(db)) as con, con:
            con.execute('CREATE VIEW billing_usage AS SELECT 1 AS x')
        with self.assertRaises(sqlite3.OperationalError):
            Store(self.settings, clock=self.clock)
        with closing(sqlite3.connect(db)) as con, con:
            self.assertEqual(con.execute('PRAGMA user_version').fetchone()[0], 2)
            self.assertEqual(con.execute('SELECT value FROM server_meta WHERE key=\'existing\'').fetchone()[0], 'safe')
            self.assertFalse(con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='host_usage'").fetchone())
        self.assertEqual(len(list((self.root/'backups').glob('usage-before-schema-v2-to-v3-*.sqlite3'))), 1)


if __name__ == '__main__':
    unittest.main()
