"""Regression checks for explicit remembered-device authentication."""
from __future__ import annotations

import hashlib
import sqlite3
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

from redink_server.config import Settings
from redink_server.store import Store, iso


class RememberTokenStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.store = Store(Settings(data_dir=self.root, credentials={}))
        self.password_hash = 'pbkdf2:test:stored-hash-1'

    def test_remember_tokens_are_random_hash_only_reusable_and_expire(self):
        token = self.store.issue_admin_remember_token(self.password_hash)
        self.assertGreaterEqual(len(token), 40)
        self.assertTrue(self.store.valid_admin_remember_token(token, self.password_hash))
        self.assertTrue(self.store.valid_admin_remember_token(token, self.password_hash))
        with self.store.connect() as conn:
            row = conn.execute('SELECT * FROM admin_remember_tokens').fetchone()
            self.assertEqual(row['token_hash'], hashlib.sha256(token.encode()).hexdigest())
            self.assertNotIn(token, tuple(row))
            self.assertEqual(len(conn.execute('SELECT * FROM admin_remember_tokens').fetchall()), 1)
            conn.execute('UPDATE admin_remember_tokens SET expires_utc=?',
                         (iso(datetime.now(timezone.utc)-timedelta(seconds=1)),))
        self.assertFalse(self.store.valid_admin_remember_token(token, self.password_hash))
        with self.store.connect() as conn:
            self.assertFalse(conn.execute('SELECT 1 FROM admin_remember_tokens').fetchone())

    def test_logout_revocation_password_change_and_unrelated_tokens(self):
        token1 = self.store.issue_admin_remember_token(self.password_hash)
        token2 = self.store.issue_admin_remember_token(self.password_hash)
        self.assertNotEqual(token1, token2)
        self.store.revoke_admin_remember_token(token1)
        self.assertFalse(self.store.valid_admin_remember_token(token1, self.password_hash))
        self.assertTrue(self.store.valid_admin_remember_token(token2, self.password_hash))
        self.assertFalse(self.store.valid_admin_remember_token(token2, 'new-password-hash'))
        self.assertFalse(self.store.valid_admin_remember_token(token2, self.password_hash))

    def test_invalid_tokens_do_not_access_storage_or_throw(self):
        for raw in ('', 'bad', 'ä'*90, 'a'*300, None):
            self.assertFalse(self.store.valid_admin_remember_token(raw, self.password_hash))
            self.store.revoke_admin_remember_token(raw)
        with self.assertRaises(ValueError):
            self.store.issue_admin_remember_token(self.password_hash, days=31)


try:
    from werkzeug.security import generate_password_hash
    from redink_server.web import create_app, REMEMBER_COOKIE
except ImportError:
    create_app = None


@unittest.skipUnless(create_app is not None, 'Flask/Werkzeug dependencies are required for HTTP integration tests')
class RememberLoginHTTPTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        credentials = {'session_secret':'s'*64,
                       'admin_password_hash':generate_password_hash('admin-password'),
                       'ingestion_token_hash':'i'*64,
                       'config_token_hash':'c'*64}
        self.settings = Settings(data_dir=self.root, credentials=credentials, secure_cookies=True)
        self.store = Store(self.settings)
        self.app = create_app(self.settings, self.store)
        self.app.testing = True
        self.client = self.app.test_client()

    def login(self, remember=False):
        self.client.get('/login')
        with self.client.session_transaction() as data:
            csrf = data['csrf']
        return self.client.post('/login', data={'csrf_token':csrf,'password':'admin-password',
                                                'remember_device':'on' if remember else ''})

    def test_optional_checkbox_and_secure_remember_cookie(self):
        page = self.client.get('/login')
        self.assertIn('Remember this browser for 30 days', page.text)
        self.assertEqual(self.login(remember=False).status_code, 302)
        self.assertNotIn(REMEMBER_COOKIE+'=', self.client.get('/healthz').headers.get('Set-Cookie',''))
        self.assertEqual(self.store.connect().__class__.__name__, '_GeneratorContextManager')

    def test_remember_login_after_session_loss_logout_revokes(self):
        response = self.login(remember=True)
        self.assertEqual(response.status_code, 302)
        cookies = response.headers.getlist('Set-Cookie')
        remember_cookie = next(v for v in cookies if v.startswith(REMEMBER_COOKIE+'='))
        self.assertIn('Secure', remember_cookie)
        self.assertIn('HttpOnly', remember_cookie)
        self.assertIn('SameSite=Strict', remember_cookie)
        self.assertIn('Max-Age=2592000', remember_cookie)
        with self.client.session_transaction() as data:
            data.clear()
        self.assertEqual(self.client.get('/api/dashboard').status_code, 200)
        with self.client.session_transaction() as data:
            csrf = data['csrf'] if 'csrf' in data else None
        if not csrf:
            self.client.get('/')
            with self.client.session_transaction() as data:
                csrf = data['csrf']
        self.assertEqual(self.client.post('/logout',data={'csrf_token':csrf}).status_code,302)
        self.assertEqual(self.client.get('/api/dashboard').status_code,401)
        with self.store.connect() as conn:
            self.assertEqual(conn.execute('SELECT count(*) FROM admin_remember_tokens').fetchone()[0],0)

    def test_token_invalidated_by_password_change(self):
        self.login(remember=True)
        self.settings.credentials['admin_password_hash']=generate_password_hash('new-password')
        with self.client.session_transaction() as data:
            data.clear()
        self.assertEqual(self.client.get('/api/dashboard').status_code,401)


if __name__ == '__main__':
    unittest.main()
