"""Listener lifecycle regression tests without a real network socket.

The mock listener uses the same prepare/serve/stop contract as Cheroot. Real
Windows/TLS socket behaviour remains an explicit platform acceptance step.
"""
from __future__ import annotations

import threading
import time
import unittest
from dataclasses import dataclass

from redink_server.listener_runtime import ListenerSupervisor


@dataclass(frozen=True)
class FakeSettings:
    host: str = '127.0.0.1'
    port: int = 8765
    tls_enabled: bool = False


class FakeListener:
    def __init__(self, settings, history, fail=False):
        self.settings = settings
        self.history = history
        self.fail = fail
        self.stop_event = threading.Event()

    def prepare(self):
        self.history.append(('prepare', self.settings.host, self.settings.port))
        if self.fail:
            raise OSError('Simulated Windows port conflict')

    def serve(self):
        self.history.append(('serve', self.settings.host, self.settings.port))
        self.stop_event.wait(5)

    def stop(self):
        self.history.append(('stop', self.settings.host, self.settings.port))
        self.stop_event.set()


class ListenerLifecycleTests(unittest.TestCase):
    def setUp(self):
        self.history = []
        self.on_applied = []
        self.on_rollback = []
        self.settings = FakeSettings()
        self.ready = threading.Event()
        self.supervisor = ListenerSupervisor(
            self.settings,
            lambda s: FakeListener(s, self.history, fail=(s.port == 8888)),
            self.on_applied.append,
            lambda s, text: self.on_rollback.append((s, text)),
            response_grace=.01)
        self.thread = threading.Thread(target=lambda: self.supervisor.run(lambda _: self.ready.set()), daemon=True)
        self.thread.start()
        self.assertTrue(self.ready.wait(3))

    def tearDown(self):
        self.supervisor.shutdown()
        self.thread.join(timeout=4)
        self.assertFalse(self.thread.is_alive(), 'Supervisor thread remained alive after shutdown')

    def until(self, condition, timeout=3):
        deadline = time.monotonic()+timeout
        while time.monotonic()<deadline:
            if condition():
                return
            time.sleep(.015)
        self.fail('Timed out waiting for asynchronous listener transition')

    def test_successful_rebind_without_process_restart(self):
        changed=FakeSettings(host='0.0.0.0', port=8765, tls_enabled=True)
        self.assertTrue(self.supervisor.reconfigure(changed, 'previous TOML'))
        self.until(lambda: self.on_applied and self.supervisor.status()['status']=='active')
        self.assertEqual(self.on_applied, [changed])
        self.assertIn(('stop','127.0.0.1',8765),self.history)
        self.assertIn(('prepare','0.0.0.0',8765),self.history)
        self.assertEqual(self.on_rollback, [])

    def test_rebind_conflict_rolls_back_previous_listener_and_config(self):
        invalid=FakeSettings(port=8888)
        self.assertTrue(self.supervisor.reconfigure(invalid, 'previous TOML'))
        self.until(lambda: self.supervisor.status()['status']=='failed')
        self.assertEqual(self.on_applied, [])
        self.assertEqual(self.on_rollback, [(self.settings, 'previous TOML')])
        self.assertEqual(self.history.count(('prepare','127.0.0.1',8765)),2)
        self.assertTrue(self.supervisor.reconfigure(FakeSettings(port=9876), 'current TOML'))
        self.until(lambda: bool(self.on_applied))
        self.assertEqual(self.on_applied[0].port,9876)

    def test_shutdown_is_deferred_outside_request_worker(self):
        self.assertTrue(self.supervisor.shutdown())
        self.assertFalse(self.supervisor.shutdown())
        self.thread.join(timeout=3)
        self.assertFalse(self.thread.is_alive())
        self.assertEqual(self.supervisor.status()['status'],'stopped')

    def test_only_one_rebind_can_be_pending(self):
        self.assertTrue(self.supervisor.reconfigure(FakeSettings(port=9877),'old'))
        self.assertFalse(self.supervisor.reconfigure(FakeSettings(port=9878),'old'))
        self.until(lambda: bool(self.on_applied))
        self.assertEqual(self.on_applied[0].port,9877)


class RuntimeWiringTests(unittest.TestCase):
    def test_listener_uses_same_transport_in_http_and_https(self):
        from pathlib import Path
        root=Path(__file__).resolve().parents[1]
        cli=(root/'redink_server/cli.py').read_text('utf-8')
        tls=(root/'redink_server/tls.py').read_text('utf-8')
        web=(root/'redink_server/web.py').read_text('utf-8')
        self.assertNotIn('from waitress import create_server',cli)
        self.assertIn('ListenerSupervisor(settings, listener_factory',cli)
        self.assertIn('create_http_server(next_settings, app)',cli)
        self.assertIn('create_https_server(next_settings, app)',cli)
        self.assertIn('from cheroot.wsgi import Server',tls)
        self.assertIn("app.extensions['redink_shutdown'] = supervisor.shutdown",cli)
        self.assertIn("supervisor.reconfigure(fresh, previous_toml)",web)
        self.assertIn('restore_configuration(settings.data_dir, previous_toml)',web)
        self.assertIn('SESSION_COOKIE_SECURE=next_settings.secure_cookies',web)
        self.assertIn('TRUSTED_HOSTS=list(next_settings.trusted_hosts)',web)

    def test_daily_chart_has_exact_values_and_accessible_table(self):
        from pathlib import Path
        root=Path(__file__).resolve().parents[1]
        js=(root/'redink_server/static/app.js').read_text('utf-8')
        html=(root/'redink_server/templates/index.html').read_text('utf-8')
        self.assertIn('id="log-chart-selection"',html)
        self.assertIn('id="log-chart-days"',html)
        self.assertIn('Show exact daily values',html)
        self.assertIn('hit.addEventListener("focus",show)',js)
        self.assertIn('d.date),element("td","numeric",formatNumber(d.standard',js)


if __name__ == '__main__':
    unittest.main()
