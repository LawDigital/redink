"""v1.3 security, configuration and disjoint user analytics regressions."""
from __future__ import annotations

import json
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from dataclasses import replace

from fixtures import usage, log_line, log_filename
from redink_server.config import ConfigError, initialize, load_settings, token_hash
from redink_server.settings_ui import (configuration_document, update_configuration, update_passwordless,
                                       is_passwordless, rotate_token, is_loopback_request, apply_live_settings)
from redink_server.protocol import normalize_event
from redink_server.scanner import Scanner
from redink_server.store import Store
from redink_server.user_analytics import (function_users, license_users, user_summary_csv, license_users_csv)


class SettingsAndUserTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.initial = initialize(self.directory)
        self.settings = load_settings(self.directory)
        self.clock = lambda: datetime(2026, 10, 31, 12, tzinfo=timezone.utc)
        self.store = Store(self.settings, clock=self.clock)

    def test_localhost_passwordless_is_encrypted_and_tampering_fails(self):
        self.assertFalse(is_passwordless(self.settings))
        updated = update_passwordless(self.settings, True)
        self.assertTrue(is_passwordless(updated))
        stored = (self.directory/'server.toml').read_text()
        self.assertNotIn('localhost-only-v1', stored)
        self.assertNotIn('passwordless_mode = true', stored)
        self.assertFalse(is_loopback_request('192.168.1.50'))
        self.assertTrue(is_loopback_request('127.0.0.1'))
        (self.directory/'server.toml').write_text(stored.replace('passwordless_mode = "', 'passwordless_mode = "tampered'))
        with self.assertRaises(ConfigError):
            load_settings(self.directory)
        (self.directory/'server.toml').write_text(stored)
        self.assertFalse(is_passwordless(update_passwordless(updated, False)))

    def test_rotation_is_independent_immediate_and_persists(self):
        old_ingestion = self.initial['ingestion_token']
        old_config = self.initial['config_token']
        old_config_hash = self.settings.credentials['config_token_hash']
        new_ingestion = rotate_token(self.settings, 'ingestion')
        self.assertNotEqual(old_ingestion, new_ingestion)
        self.assertNotEqual(token_hash(old_ingestion), self.settings.credentials['ingestion_token_hash'])
        self.assertEqual(token_hash(new_ingestion), self.settings.credentials['ingestion_token_hash'])
        self.assertEqual(old_config_hash, self.settings.credentials['config_token_hash'])
        persisted = load_settings(self.directory)
        self.assertEqual(persisted.credentials['ingestion_token_hash'], token_hash(new_ingestion))
        document = configuration_document(persisted)
        self.assertIn('publishing', document['sections'])
        newer_config = rotate_token(self.settings, 'config')
        self.assertNotEqual(newer_config, old_config)
        self.assertEqual(token_hash(new_ingestion), self.settings.credentials['ingestion_token_hash'])
        with self.assertRaises(ConfigError):
            rotate_token(self.settings, 'admin')

    def test_settings_validation_preserves_original_config(self):
        original = (self.directory/'server.toml').read_text()
        with self.assertRaises(ConfigError):
            update_configuration(self.settings, {'server': {'port': 90000}})
        self.assertEqual((self.directory/'server.toml').read_text(), original)
        with self.assertRaises(ConfigError):
            update_configuration(self.settings, {'authentication': {'passwordless_mode': 'true'}})
        result = update_configuration(self.settings, {'imports': {'log_paths': ['a','b']},
                                                      'publishing': {'source_directory': 'shared'},
                                                      'billing': {'currency': 'EUR'}})
        self.assertFalse(result['restart_required'])
        loaded = load_settings(self.directory)
        self.assertEqual(len(loaded.log_paths), 2)
        self.assertEqual(loaded.currency, 'EUR')
        self.assertTrue(str(loaded.config_directory).endswith('shared'))

    def test_live_settings_apply_to_importer_and_publishing_without_restart(self):
        source = self.directory/'live-source'
        source.mkdir()
        ini = self.directory/'approved-ini'
        ini.mkdir()
        (ini/'redink.ini').write_text('RedInk=true\n', encoding='utf-8')
        path = source/log_filename('alice','EXAMPLE','WD')
        path.write_text(log_line('RI_QuickTranslate_Click invoked',user='alice',timestamp='2026-10-31 08:00:00.000'))
        changed = {'imports': {'log_paths': [str(source)],'scan_seconds': 10},
                   'publishing': {'source_directory': str(ini)},
                   'billing': {'currency': 'EUR'}}
        result = update_configuration(self.settings, changed)
        self.assertFalse(result['restart_required'])
        validated = load_settings(self.directory)
        active, applied = apply_live_settings(self.settings, validated, changed)
        self.assertNotEqual(active, self.settings)
        self.assertEqual(active.config_directory, ini)
        self.assertEqual(active.currency, 'EUR')
        self.assertEqual(active.scan_seconds, 10)
        self.assertIn('log_paths', applied)
        scanner = Scanner(self.settings, self.store)
        scanner.update_settings(active)
        self.store.settings=active
        self.assertTrue(scanner._wake.is_set())
        results = scanner.scan_all()
        self.assertTrue(any(record['kind']=='log' and record['records']==1 for record in results))
        scanner.stop()

    def test_listener_change_remains_restart_only_when_live_fields_also_change(self):
        changed = {'server': {'port': 9876},'imports': {'log_paths':['changed']}}
        result = update_configuration(self.settings,changed)
        self.assertTrue(result['restart_required'])
        loaded=load_settings(self.directory)
        active,applied=apply_live_settings(self.settings,loaded,changed)
        self.assertEqual(active.port,self.settings.port)
        self.assertEqual(loaded.port,9876)
        self.assertEqual(len(active.log_paths),1)
        self.assertIn('log_paths',applied)

    def test_user_stats_and_disjoint_licence_identity(self):
        folder=self.directory/'incoming'
        folder.mkdir()
        path=folder/log_filename('alice','EXAMPLE','WD')
        path.write_text(log_line('AgentRun_Word_Agent invoked',user='alice',timestamp='2026-10-30 08:00:00.000')+
                        log_line('RI_QuickTranslate_Click invoked',user='alice',timestamp='2026-10-31 08:00:00.000')+
                        log_line('RI_QuickTranslate_Click invoked',user='alice',timestamp='2026-10-31 09:00:00.000'))
        scanner=Scanner(self.settings,self.store)
        scanner.scan_path(folder,'log')
        data=function_users(self.store,'30d','WD')
        self.assertEqual(data['total_invocations'],3)
        self.assertEqual(data['total_users'],1)
        self.assertEqual(data['users'][0]['active_days'],2)
        self.assertEqual(data['users'][0]['agentic'],1)
        detail=function_users(self.store,'30d','WD',data['users'][0]['user_hash'])['selected']
        self.assertEqual(detail['standard'],2)
        self.assertIn('Agent run starts',user_summary_csv(data))
        self.store.ingest(normalize_event(usage(loaded_dates=['2026-10-30'],loaded_hosts=['WD'])))
        license_data=license_users(self.store,'2026-10')
        self.assertEqual(len(license_data['users']),1)
        self.assertNotEqual(data['users'][0]['user_hash'],license_data['users'][0]['user_key'])
        self.assertIn('User key',license_users_csv(license_data))
        scanner.stop()

if __name__ == '__main__':
    unittest.main()
