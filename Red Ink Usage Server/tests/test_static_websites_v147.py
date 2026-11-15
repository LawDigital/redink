"""Static hosting regression tests; no real network or Flask server required.

Integration with Windows Cheroot/TLS is a separate real-machine smoke test.
"""
from __future__ import annotations

import json
import tempfile
import threading
import time
import unittest
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

from cryptography.fernet import Fernet

from redink_server.config import ConfigError, DEFAULT_TOML, StaticWebsite, Settings, load_settings
from redink_server.settings_ui import (replace_values, update_configuration,
                                        restore_configuration, update_passwordless)
from redink_server.static_websites import (WebsiteListenerGroup, create_all_listeners,
                                           resolve_static_file)


def entry(name, directory, *, port=9443, path='/', host='127.0.0.1', https=False,
          enabled=True, spa_fallback=False):
    return dict(name=name, enabled=enabled, host=host, port=port, path=path,
                directory=str(directory), https=https, spa_fallback=spa_fallback)


class StaticWebConfigurationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.data = self.root / 'data'
        self.data.mkdir()
        self.appdir = self.root / 'public-app'
        self.appdir.mkdir()
        (self.appdir / 'index.html').write_text('<h1>Intranet</h1>')
        (self.data / 'server.toml').write_bytes(DEFAULT_TOML.replace('\n', '\r\n').encode('utf-8'))
        (self.data / 'secrets.json').write_text(json.dumps(dict(
            session_secret='a'*64, admin_password_hash='b'*64,
            ingestion_token_hash='c'*64, config_token_hash='d'*64,
            configuration_encryption_key=Fernet.generate_key().decode('ascii'))))
        self.settings = load_settings(self.data)

    def test_existing_configuration_without_websites_remains_compatible(self):
        self.assertEqual(self.settings.websites, ())
        self.assertIn('websites', DEFAULT_TOML)

    def test_sites_on_separate_ports_and_shared_paths_roundtrip_with_comments(self):
        # Two paths share port 9443; another site uses port 9883.
        sites = [entry('Portal', self.appdir, path='/apps'),
                 entry('Downloads', self.appdir, path='/downloads'),
                 entry('Other', self.appdir, port=9883, path='/')]
        first = update_configuration(self.settings, {'websites': {'sites': sites}})
        self.assertTrue(first['restart_required'])
        loaded = load_settings(self.data)
        self.assertEqual([x.name for x in loaded.websites], ['Portal', 'Downloads', 'Other'])
        self.assertEqual([x.path for x in loaded.websites], ['/apps', '/downloads', '/'])
        second = update_configuration(loaded, {'websites': {'sites': [sites[0]]}})
        self.assertTrue(second['saved'])
        self.assertEqual([x.name for x in load_settings(self.data).websites], ['Portal'])
        self.assertIn('# Optional independent read-only static websites', (self.data/'server.toml').read_text())

    def test_invalid_updates_roll_back_original_toml(self):
        original = (self.data/'server.toml').read_bytes()
        violations = [
            [entry('Exposes secrets', self.data)],
            [entry('Root contains secrets', self.root)],
            [entry('Hidden', self.appdir, path='/../other')],
            [entry('Invalid port', self.appdir, port=90000)],
            [entry('Admin root', self.appdir, port=8765, path='/')],
            [entry('Duplicate', self.appdir), entry('Duplicate', self.appdir, port=9881)],
            [entry('Same prefix', self.appdir), entry('Other', self.appdir)],
            [entry('Overlapping bind', self.appdir, host='0.0.0.0'),
             entry('Other', self.appdir, host='127.0.0.1', path='/other')],
            [entry('Same loopback', self.appdir, host='localhost'),
             entry('Other', self.appdir, host='127.0.0.1', path='/other')],
            [entry('No TLS', self.appdir, https=True)],
            [entry('Plain remote', self.appdir, host='0.0.0.0')],
            [entry('Missing folder', self.root/'missing')],
        ]
        for case in violations:
            with self.subTest(case=case):
                with self.assertRaises(ConfigError):
                    update_configuration(self.settings, {'websites': {'sites': case}})
                self.assertEqual((self.data/'server.toml').read_bytes(), original)
        with self.assertRaises(ConfigError):
            update_configuration(self.settings, {'websites': {'sites': [{'name':'Incomplete'}]}})
        self.assertEqual((self.data/'server.toml').read_bytes(), original)

    def test_valid_updates_keep_original_crlf_and_unedited_comments(self):
        path = self.data / 'server.toml'
        original = path.read_bytes()
        self.assertIn(b'\r\n', original)
        update_configuration(self.settings, {'server': {'threads': 7},
                                             'websites': {'sites': [entry('Portal', self.appdir)]}})
        changed = path.read_bytes()
        self.assertNotIn(b'\n', changed.replace(b'\r\n', b''))
        self.assertIn(b'# Relative paths below are relative to this data directory.\r\n', changed)
        self.assertEqual(load_settings(self.data).threads, 7)
        self.assertEqual(len(load_settings(self.data).websites), 1)
        restore_configuration(self.data, original.decode('utf-8'))
        self.assertEqual(path.read_bytes(), original)

    def test_invalid_config_restores_crlf_exactly_after_validation_failure(self):
        # Force failure *after* atomic replacement, not only during parsing.
        path = self.data / 'server.toml'
        original = path.read_bytes()
        with self.assertRaises(ConfigError):
            update_configuration(self.settings, {'server': {'port': 90000}})
        self.assertEqual(path.read_bytes(), original)
        with patch('redink_server.settings_ui.load_settings', side_effect=ConfigError('invalid after save')):
            with self.assertRaises(ConfigError):
                update_configuration(self.settings, {'server': {'port': 9999}})
        self.assertEqual(path.read_bytes(), original)

    def test_passwordless_setting_preserves_crlf_on_success_and_rollback(self):
        path = self.data / 'server.toml'
        original = path.read_bytes()
        update_passwordless(self.settings, True)
        saved = path.read_bytes()
        self.assertNotIn(b'\n', saved.replace(b'\r\n', b''))
        self.assertIn(b'passwordless_mode = "gAAAA', saved)
        restore_configuration(self.data, original.decode('utf-8'))
        self.assertEqual(path.read_bytes(), original)
        with patch('redink_server.settings_ui.load_settings', side_effect=ConfigError('invalid after save')):
            with self.assertRaises(ConfigError):
                update_passwordless(self.settings, True)
        self.assertEqual(path.read_bytes(), original)

    def test_lf_only_configuration_is_not_rewritten_to_crlf(self):
        path = self.data / 'server.toml'
        path.write_bytes(DEFAULT_TOML.encode('utf-8'))
        update_configuration(self.settings, {'server': {'threads': 6}})
        changed = path.read_bytes()
        self.assertNotIn(b'\r\n', changed)
        self.assertIn(b'threads = 6\n', changed)

    def test_admin_port_allows_distinct_site_prefixes_without_rebinding(self):
        sites = [entry('Portal', self.appdir, port=8765, path='/apps'),
                 entry('Reports', self.appdir, port=8765, path='/reports-app')]
        update_configuration(self.settings, {'websites': {'sites': sites}})
        loaded = load_settings(self.data)
        self.assertEqual([site.path for site in loaded.websites], ['/apps', '/reports-app'])
        self.assertEqual([site.port for site in loaded.websites], [8765, 8765])

    def test_admin_https_port_can_share_one_certificate_and_wildcard_bind(self):
        site = entry('LIL Apps', self.appdir, port=8765, path='/apps',
                     host='0.0.0.0', https=True)
        update_configuration(self.settings, {'server': {'host': '0.0.0.0'},
                                             'tls': {'enabled': True,
                                                     'certificate': 'tls/server.pem',
                                                     'private_key': 'tls/key.pem'},
                                             'websites': {'sites': [site]}})
        loaded = load_settings(self.data)
        self.assertTrue(loaded.tls_enabled)
        self.assertEqual(loaded.host, '0.0.0.0')
        self.assertEqual(len(loaded.websites), 1)
        self.assertEqual(loaded.websites[0].port, loaded.port)
        class AdminApp:
            class Map:
                @staticmethod
                def iter_rules():
                    return [type('Rule', (), {'rule': '/api/settings'})()]
            url_map = Map()
            def __call__(self, environ, respond):
                respond('200 OK', [])
                return [b'ADMIN']
        made = []
        def make_https(settings, app):
            made.append((settings.host, settings.port, app))
            return StubListener(settings.host, settings.port)
        def static_app(environ, respond):
            respond('200 OK', [])
            return [b'STATIC']
        with patch('redink_server.static_websites.create_https_server', side_effect=make_https), \
             patch('redink_server.static_websites.create_static_application', return_value=static_app):
            group = create_all_listeners(loaded, AdminApp())
        self.assertEqual(len(group._listeners), 1)
        self.assertEqual(len(made), 1)
        self.assertEqual(made[0][:2], ('0.0.0.0', 8765))
        self.assertEqual(b''.join(made[0][2]({'PATH_INFO':'/apps/'}, lambda *args:None)), b'STATIC')
        self.assertEqual(b''.join(made[0][2]({'PATH_INFO':'/api/settings'}, lambda *args:None)), b'ADMIN')

    def test_admin_port_cannot_override_protected_usage_routes_or_other_bind(self):
        invalid = [
            entry('Root', self.appdir, port=8765, path='/'),
            entry('API', self.appdir, port=8765, path='/api/apps'),
            entry('Collector', self.appdir, port=8765, path='/collect/data'),
            entry('Login', self.appdir, port=8765, path='/login'),
            entry('Static', self.appdir, port=8765, path='/static'),
            entry('Other host', self.appdir, port=8765, path='/apps',host='0.0.0.0'),
            entry('Other scheme', self.appdir, port=8765, path='/apps',https=True),
        ]
        original = (self.data/'server.toml').read_bytes()
        for site in invalid:
            with self.subTest(site=site['name']):
                with self.assertRaises(ConfigError):
                    update_configuration(self.settings, {'websites': {'sites': [site]}})
                self.assertEqual((self.data/'server.toml').read_bytes(), original)

    def test_enabled_https_site_reuses_existing_tls_setting(self):
        # Enabling main TLS is sufficient for subordinate sites in config;
        # certificate validity is separately checked by the existing TLS code.
        from redink_server.settings_ui import update_configuration
        site = entry('Portal', self.appdir, host='0.0.0.0', https=True)
        changed = {'tls': {'enabled': True, 'certificate': 'tls/cert.pem',
                           'private_key': 'tls/key.pem'}, 'websites': {'sites': [site]}}
        update_configuration(self.settings, changed)
        loaded = load_settings(self.data)
        self.assertTrue(loaded.websites[0].https)
        self.assertTrue(loaded.tls_enabled)

    def test_site_directory_is_absolute_and_never_descends_from_private_data(self):
        with self.assertRaises(ConfigError):
            update_configuration(self.settings, {'websites': {'sites': [entry('Bad', 'relative/path')]}})


class StaticWebFileTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root/'index.html').write_text('Home')
        (self.root/'script.js').write_text('console.log(1)')
        (self.root/'app.json').write_text('{}')
        (self.root/'.env').write_text('TOKEN=secret')
        (self.root/'server.toml').write_text('private')
        (self.root/'docs').mkdir()
        (self.root/'docs'/'index.htm').write_text('Docs')
        (self.root/'empty').mkdir()
        self.site = StaticWebsite('Portal', True, '127.0.0.1', 9443, '/apps', self.root, False)

    def test_root_and_nested_index_and_regular_assets(self):
        self.assertEqual(resolve_static_file(self.site, '', True)[0], self.root/'index.html')
        self.assertEqual(resolve_static_file(self.site, 'docs/', True)[0], self.root/'docs'/'index.htm')
        self.assertEqual(resolve_static_file(self.site, 'script.js', False)[0], self.root/'script.js')
        self.assertEqual(resolve_static_file(self.site, 'app.json', False)[0], self.root/'app.json')
        self.assertEqual(resolve_static_file(self.site, 'docs', False), (None, True))

    def test_no_directory_listing_or_secret_access(self):
        for name in ('empty/', '.env', '.git/config', 'server.toml',
                     '../elsewhere', '..', 'docs/../index.html', 'docs\\index.htm',
                     'README.pem', 'backend.php', 'nonexistent.js', 'some:ads', 'docs//index.htm'):
            with self.subTest(name=name):
                self.assertEqual(resolve_static_file(self.site, name, True)[0], None)
        # Symlinks must never escape the configured document root.
        outside = self.root.parent/'outside-private-file.txt'
        outside.write_text('private')
        self.addCleanup(lambda: outside.unlink(missing_ok=True))
        try:
            (self.root/'outside.txt').symlink_to(outside)
        except (OSError, NotImplementedError):
            return
        self.assertEqual(resolve_static_file(self.site, 'outside.txt', False)[0], None)

    def test_spa_fallback_is_optional_and_only_for_routes_not_files(self):
        self.assertEqual(resolve_static_file(self.site, 'dashboard/users', False)[0], None)
        spa = replace(self.site, spa_fallback=True)
        self.assertEqual(resolve_static_file(spa, 'dashboard/users', False)[0], self.root/'index.html')
        self.assertEqual(resolve_static_file(spa, 'missing.js', False)[0], None)
        self.assertEqual(resolve_static_file(spa, '.env', False)[0], None)


class StubListener:
    def __init__(self, host, port, *, fail=False):
        self.host, self.port, self.fail = host, port, fail
        self.prepared = False
        self.stopped = 0
        self.end = threading.Event()

    def prepare(self):
        if self.fail:
            raise OSError('Port in use')
        self.prepared = True

    def serve(self):
        self.end.wait(3)

    def stop(self):
        self.stopped += 1
        self.end.set()


class StaticListenerTests(unittest.TestCase):
    def test_port_conflict_stops_every_prepared_listener(self):
        first, second = StubListener('127.0.0.1',8765), StubListener('127.0.0.1',9443,fail=True)
        with self.assertRaises(OSError):
            WebsiteListenerGroup([first, second]).prepare()
        self.assertEqual(first.stopped,1)
        self.assertFalse(second.prepared)

    def test_listener_group_is_stoppable_and_joins_all_workers(self):
        listeners = [StubListener('127.0.0.1',8765),StubListener('127.0.0.1',9443)]
        group = WebsiteListenerGroup(listeners)
        group.prepare()
        thread=threading.Thread(target=group.serve)
        thread.start()
        time.sleep(.05)
        group.stop()
        thread.join(2)
        self.assertFalse(thread.is_alive())
        self.assertEqual([x.stopped for x in listeners],[1,1])

    def test_failed_additional_site_rolls_back_entire_listener_generation(self):
        from redink_server.listener_runtime import ListenerSupervisor
        initial = object()
        changed = object()
        restored = []
        applied = []
        ready = threading.Event()

        def factory(settings):
            return WebsiteListenerGroup([
                StubListener('127.0.0.1', 8765),
                StubListener('127.0.0.1', 9443, fail=(settings is changed))])

        supervisor = ListenerSupervisor(initial, factory, applied.append,
                                        lambda previous, original: restored.append((previous, original)),
                                        response_grace=.01)
        worker = threading.Thread(target=lambda: supervisor.run(lambda _: ready.set()))
        worker.start()
        self.assertTrue(ready.wait(3))
        try:
            self.assertTrue(supervisor.reconfigure(changed, 'prior TOML'))
            for _ in range(200):
                if supervisor.status()['status'] == 'failed':
                    break
                time.sleep(.01)
            self.assertEqual(supervisor.status()['status'], 'failed')
            self.assertEqual(restored, [(initial, 'prior TOML')])
            self.assertEqual(applied, [])
        finally:
            supervisor.shutdown()
            worker.join(4)
            self.assertFalse(worker.is_alive())

    def test_website_ports_are_independent_and_paths_can_share_one(self):
        from redink_server.config import Settings
        settings=Settings(data_dir=Path('/tmp/private'), credentials={'secret':'secret'}, websites=(
            StaticWebsite('A',True,'127.0.0.1',9443,'/apps',Path('/tmp/a'),False),
            StaticWebsite('B',True,'127.0.0.1',9443,'/downloads',Path('/tmp/b'),False),
            StaticWebsite('C',True,'127.0.0.1',9555,'/',Path('/tmp/c'),False)))
        made=[]
        def factory(s, app):
            made.append((s.host,s.port,app))
            return StubListener(s.host,s.port)
        with patch('redink_server.static_websites.create_http_server',side_effect=factory), \
             patch('redink_server.static_websites.create_static_application',side_effect=lambda sites: tuple(sites)):
            group=create_all_listeners(settings,'private admin app')
        self.assertEqual(len(made),3)
        self.assertEqual(made[0][2],'private admin app')
        self.assertEqual([site.path for site in made[1][2]],['/apps','/downloads'])
        self.assertEqual(made[2][1],9555)
        self.assertEqual(len(group._listeners),3)

    def test_admin_and_sites_share_one_wsgi_listener_without_route_leaks(self):
        from types import SimpleNamespace
        from redink_server.static_websites import create_shared_application
        admin_calls, site_calls, factories = [], [], []
        class AdminApp:
            url_map = SimpleNamespace(iter_rules=lambda: [SimpleNamespace(rule=v) for v in
                ('/', '/login', '/api/collect', '/collect/<token>', '/custom-secure/action')])
            def __call__(self, environ, respond):
                admin_calls.append(environ['PATH_INFO'])
                respond('200 OK', [])
                return [b'ADMIN']
        admin = AdminApp()
        def static_app(environ, respond):
            site_calls.append(environ['PATH_INFO'])
            respond('200 OK', [])
            return [b'STATIC']
        settings = Settings(data_dir=Path('/tmp/private'), credentials={}, websites=(
            StaticWebsite('Portal',True,'127.0.0.1',8765,'/apps',Path('/tmp/a'),False),
            StaticWebsite('Docs',True,'127.0.0.1',8765,'/docs',Path('/tmp/b'),False),
            StaticWebsite('External',True,'127.0.0.1',9443,'/',Path('/tmp/c'),False)))
        def make_listener(s, app):
            factories.append((s.port,app))
            return StubListener(s.host,s.port)
        with patch('redink_server.static_websites.create_http_server', side_effect=make_listener), \
             patch('redink_server.static_websites.create_static_application',return_value=static_app):
            group = create_all_listeners(settings,admin)
        self.assertEqual(len(group._listeners),2)
        self.assertEqual([entry[0] for entry in factories],[8765,9443])
        router=factories[0][1]
        for path,expected in (('/', b'ADMIN'),('/login',b'ADMIN'),('/api/collect',b'ADMIN'),
                              ('/collect/secret',b'ADMIN'),('/custom-secure/action',b'ADMIN'),
                              ('/apps',b'STATIC'),('/apps/index.html',b'STATIC'),
                              ('/docs/',b'STATIC'),('/apps-evil',b'ADMIN')):
            with self.subTest(path=path):
                self.assertEqual(b''.join(router({'PATH_INFO':path},lambda *args:None)),expected)
        self.assertIn('/collect/secret',admin_calls)
        self.assertEqual(site_calls,['/apps','/apps/index.html','/docs/'])
        with patch('redink_server.static_websites.create_static_application',return_value=static_app):
            with self.assertRaises(ConfigError):
                create_shared_application(admin,(StaticWebsite('Blocked',True,'127.0.0.1',8765,
                    '/custom-secure',Path('/tmp/a'),False),))



if __name__ == '__main__':
    unittest.main()
