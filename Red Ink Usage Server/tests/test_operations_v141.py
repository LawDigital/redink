"""Security and user-interface regressions for 1.4.1 server operations.

Only static contracts run on non-Windows: full browser and Task Scheduler
integration are mandatory release acceptance checks on Windows.
"""
from pathlib import Path
import os
import unittest

from redink_server.startup import StartupError, manage_autostart

ROOT = Path(__file__).resolve().parents[1]


class OperationsContractTests(unittest.TestCase):
    def test_autostart_rejects_unrecognised_actions(self):
        with self.assertRaises(StartupError):
            manage_autostart(ROOT, 'delete-everything')

    @unittest.skipIf(os.name == 'nt', 'Non-Windows rejection only')
    def test_autostart_rejects_nonwindows(self):
        with self.assertRaisesRegex(StartupError, 'Windows'):
            manage_autostart(ROOT, 'enable')

    def test_privileged_operations_are_csrf_and_localhost_only(self):
        source = (ROOT / 'redink_server/web.py').read_text(encoding='utf-8')
        for endpoint in ("@app.post('/api/settings/shutdown')", "@app.post('/api/settings/autostart')"):
            self.assertIn(endpoint, source)
        for action in ('def server_shutdown():', 'def autostart_change():'):
            body = source.split(action, 1)[1].split('\n    @app.', 1)[0]
            self.assertIn('check_csrf()', body)
            self.assertIn('is_loopback_request(request.remote_addr)', body)
        self.assertIn("app.extensions.get('redink_shutdown')", source)

    def test_urls_are_complete_and_copyable_without_hardcoded_tokens(self):
        js = (ROOT / 'redink_server/static/app.js').read_text(encoding='utf-8')
        template = (ROOT / 'redink_server/templates/index.html').read_text(encoding='utf-8')
        for name in ('redink.ini', 'allmodels.ini', 'specialservices.ini'):
            self.assertIn(name, js)
        self.assertIn('"/config/"+encodeURIComponent(configToken)', js)
        self.assertIn('"/collect/"+encodeURIComponent(ingestionToken)', js)
        self.assertIn('new URL(path,location.origin).href', js)
        self.assertIn('const base = location.origin;', js)
        self.assertIn('Copy link', js)
        self.assertIn('id="client-url-list"', template)
        self.assertIn('id="shutdown-server"', template)
        self.assertIn('id="autostart-enable"', template)

    def test_current_user_script_is_headless_and_not_a_service_task(self):
        source = (ROOT / 'scripts/user-autostart.ps1').read_text(encoding='utf-8')
        self.assertIn('New-ScheduledTaskTrigger -AtLogOn -User', source)
        self.assertIn('-LogonType Interactive -RunLevel Limited', source)
        self.assertIn('serve --no-browser --data-dir', source)
        self.assertIn('Unregister-ScheduledTask', source)

    def test_admin_guide_shipped_with_docs(self):
        content = (ROOT/'docs/ADMINISTRATOR_GUIDE.md').read_text(encoding='utf-8')
        for term in ('LicenseCounterPath', 'LicenseCounterIntervall', 'LicenseCounterAnon',
                     'LogPath', 'Windows login autostart', 'License scope',
                     'redink.ini.sig', 'Shut down server', 'Token rotation'):
            self.assertIn(term.lower(), content.lower())


if __name__ == '__main__':
    unittest.main()
