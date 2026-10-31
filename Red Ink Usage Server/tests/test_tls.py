"""Certificate checks and real HTTPS requests using a local, trusted test CA."""
import errno
import ipaddress
import io
import json
import os
import socket
import ssl
import tempfile
import threading
import unittest
from dataclasses import replace
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch
from urllib.error import URLError
from urllib.request import urlopen

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID
from flask import Flask, jsonify, request

from redink_server.config import ConfigError, Settings, initialize, load_settings
from redink_server.cli import main
from redink_server.tls import create_https_server, import_pfx, validate_tls


class TlsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.now = datetime.now(timezone.utc)
        cls.ca_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "Red Ink test CA")])
        cls.ca = (x509.CertificateBuilder().subject_name(name).issuer_name(name)
                  .public_key(cls.ca_key.public_key()).serial_number(x509.random_serial_number())
                  .not_valid_before(cls.now - timedelta(days=1)).not_valid_after(cls.now + timedelta(days=30))
                  .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
                  .add_extension(x509.KeyUsage(digital_signature=True, content_commitment=False, key_encipherment=False,
                    data_encipherment=False, key_agreement=False, key_cert_sign=True, crl_sign=True,
                    encipher_only=False, decipher_only=False), critical=True)
                  .add_extension(x509.SubjectKeyIdentifier.from_public_key(cls.ca_key.public_key()), critical=False)
                  .sign(cls.ca_key, hashes.SHA256()))
        cls.key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        cls.certificate = cls._leaf(cls.now - timedelta(hours=1), cls.now + timedelta(days=10))

    @classmethod
    def _leaf(cls, start, end):
        name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")])
        return (x509.CertificateBuilder().subject_name(name).issuer_name(cls.ca.subject)
                .public_key(cls.key.public_key()).serial_number(x509.random_serial_number())
                .not_valid_before(start).not_valid_after(end)
                .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
                .add_extension(x509.SubjectAlternativeName([x509.DNSName("localhost"), x509.IPAddress(ipaddress.ip_address("127.0.0.1"))]), critical=False)
                .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
                .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(cls.ca_key.public_key()), critical=False)
                .sign(cls.ca_key, hashes.SHA256()))

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.cert = self.root / "fullchain.pem"
        self.key_path = self.root / "private-key.pem"
        self.ca_path = self.root / "ca.pem"
        self.ca_path.write_bytes(self.ca.public_bytes(serialization.Encoding.PEM))
        self.cert.write_bytes(self.certificate.public_bytes(serialization.Encoding.PEM) + self.ca.public_bytes(serialization.Encoding.PEM))
        self.key_path.write_bytes(self.key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
        self.settings = Settings(self.root, {}, tls_enabled=True, tls_certificate=self.cert, tls_private_key=self.key_path, secure_cookies=True)

    def test_localhost_http_needs_no_certificate_and_old_config_still_loads(self):
        data = self.root / "plain"
        initialize(data)
        settings = load_settings(data)
        self.assertEqual(settings.host, "127.0.0.1")
        self.assertFalse(settings.tls_enabled)
        self.assertFalse(settings.secure_cookies)
        self.assertEqual(validate_tls(settings), {"enabled": False})
        path = data / "server.toml"
        contents = path.read_text()
        start = contents.index("[tls]")
        end = contents.index("[collection]")
        path.write_text(contents[:start] + contents[end:])
        self.assertFalse(load_settings(data).tls_enabled)

    def test_configured_https_paths_and_secure_cookies(self):
        data = self.root / "configured"
        initialize(data)
        path = data / "server.toml"
        text = path.read_text().replace('enabled = false', 'enabled = true').replace('certificate = ""', 'certificate = "tls/fullchain.pem"').replace('private_key = ""', 'private_key = "tls/key.pem"')
        path.write_text(text)
        settings = load_settings(data)
        self.assertTrue(settings.tls_enabled)
        self.assertTrue(settings.secure_cookies)
        self.assertEqual(settings.tls_certificate, data / "tls/fullchain.pem")
        path.write_text(text.replace('key_password_env = ""', 'key_password_env = "not an environment variable"'))
        with self.assertRaises(ConfigError):
            load_settings(data)

    def test_valid_certificate_and_encrypted_key(self):
        metadata = validate_tls(self.settings)
        self.assertIn("localhost", metadata["dns_names"])
        self.assertIn("127.0.0.1", metadata["ip_addresses"])
        password = "test-private-password"
        self.key_path.write_bytes(self.key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8,
                                                        serialization.BestAvailableEncryption(password.encode())))
        settings = replace(self.settings, tls_key_password_env="REDINK_TEST_TLS_PASSWORD")
        with patch.dict(os.environ, {"REDINK_TEST_TLS_PASSWORD": password}):
            self.assertTrue(validate_tls(settings)["enabled"])
        with patch.dict(os.environ, {"REDINK_TEST_TLS_PASSWORD": "wrong password"}):
            with self.assertRaises(ConfigError):
                validate_tls(settings)
        with self.assertRaises(ConfigError):
            validate_tls(self.settings)  # No terminal prompt/fallback for an encrypted key.

    def test_bad_key_or_certificate_cannot_start_plain_http(self):
        self.key_path.write_bytes(self.ca_key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
        with self.assertRaises(ConfigError):
            create_https_server(self.settings, lambda env, start: [])
        self.cert.write_text("not a certificate")
        with self.assertRaises(ConfigError):
            validate_tls(self.settings)

    def test_expired_or_future_certificate_is_rejected(self):
        for cert in (self._leaf(self.now - timedelta(days=10), self.now - timedelta(days=1)),
                     self._leaf(self.now + timedelta(days=1), self.now + timedelta(days=2))):
            self.cert.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
            with self.assertRaises(ConfigError):
                validate_tls(self.settings)

    def test_private_key_cannot_be_in_published_configuration_directory(self):
        published = self.root / "config-files"
        published.mkdir()
        key = published / "redink.ini"
        key.write_bytes(self.key_path.read_bytes())
        with self.assertRaises(ConfigError):
            validate_tls(replace(self.settings, tls_private_key=key))

    def test_pfx_import_is_encrypted_and_never_overwrites(self):
        source = self.root / "source.pfx"
        source.write_bytes(pkcs12.serialize_key_and_certificates(b"test", self.key, self.certificate, [self.ca],
                            serialization.BestAvailableEncryption(b"test-pfx-password")))
        result = import_pfx(self.root, source, "test-pfx-password", "test-export-password", name="renewal-2026")
        self.assertIn(b"ENCRYPTED PRIVATE KEY", Path(result["private_key"]).read_bytes())
        with patch.dict(os.environ, {"REDINK_TEST_TLS_PASSWORD": "test-export-password"}):
            settings = replace(self.settings, tls_certificate=Path(result["certificate"]), tls_private_key=Path(result["private_key"]), tls_key_password_env="REDINK_TEST_TLS_PASSWORD")
            self.assertTrue(validate_tls(settings)["enabled"])
        original = Path(result["private_key"]).read_bytes()
        with self.assertRaises(ConfigError):
            import_pfx(self.root, source, "test-pfx-password", "test-export-password", name="renewal-2026")
        self.assertEqual(original, Path(result["private_key"]).read_bytes())
        with self.assertRaises(ConfigError):
            import_pfx(self.root, source, "wrong", "test-export-password", name="other")
        self.assertFalse((self.root / "tls/other").exists())

    def test_pfx_cli_prints_only_config_paths_and_environment_names(self):
        data = self.root / "cli-data"
        initialize(data)
        source = self.root / "cli.pfx"
        source.write_bytes(pkcs12.serialize_key_and_certificates(b"test", self.key, self.certificate, [self.ca],
                            serialization.BestAvailableEncryption(b"test-pfx-password")))
        output, errors = io.StringIO(), io.StringIO()
        with patch.dict(os.environ, {"REDINK_PFX_TEST": "test-pfx-password", "REDINK_KEY_TEST": "test-export-password"}), redirect_stdout(output), redirect_stderr(errors):
            result = main(["import-pfx", "--data-dir", str(data), "--source", str(source), "--name", "test-2026",
                           "--pfx-password-env", "REDINK_PFX_TEST", "--key-password-env", "REDINK_KEY_TEST"])
        self.assertEqual(result, 0, errors.getvalue())
        self.assertIn('certificate = "tls/test-2026/fullchain.pem"', output.getvalue())
        self.assertIn('key_password_env = "REDINK_KEY_TEST"', output.getvalue())
        self.assertNotIn("test-pfx-password", output.getvalue() + errors.getvalue())
        self.assertNotIn("test-export-password", output.getvalue() + errors.getvalue())
        self.assertFalse(load_settings(data).tls_enabled)  # Activation is explicit.

    def _start_https(self, settings=None):
        app = Flask("tls-test")
        @app.get("/healthz")
        def health():
            return jsonify(status="ok", scheme=request.scheme)
        server = create_https_server(replace(settings or self.settings, port=0), app)
        try:
            server.prepare()
        except OSError as exc:
            if exc.errno == errno.EPERM or ("No socket could be created" in str(exc) and "[Errno 1] Operation not permitted" in str(exc)):
                self.skipTest("Local sockets are restricted; run the TLS integration test with local network permission.")
            raise
        thread = threading.Thread(target=server.serve, daemon=True)
        thread.start()
        self.addCleanup(lambda: thread.join(timeout=5))
        self.addCleanup(server.stop)
        return server.socket.getsockname()[1]

    def test_real_https_request_verifies_ca_and_hostname(self):
        port = self._start_https()
        context = ssl.create_default_context(cafile=self.ca_path)
        with urlopen(f"https://127.0.0.1:{port}/healthz", context=context, timeout=5) as response:
            self.assertEqual(json.load(response), {"status": "ok", "scheme": "https"})

    def test_untrusted_certificate_is_rejected_by_client(self):
        port = self._start_https()
        with self.assertRaises(URLError):
            urlopen(f"https://127.0.0.1:{port}/healthz", context=ssl.create_default_context(), timeout=5)

    def test_hostname_mismatch_is_rejected_by_client(self):
        port = self._start_https()
        context = ssl.create_default_context(cafile=self.ca_path)
        with socket.create_connection(("127.0.0.1", port), timeout=5) as connection:
            with self.assertRaises(ssl.SSLCertVerificationError):
                context.wrap_socket(connection, server_hostname="other.invalid")

    def test_encrypted_pem_key_serves_real_https(self):
        password = "native-https-test-password"
        self.key_path.write_bytes(self.key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8,
                                                       serialization.BestAvailableEncryption(password.encode())))
        with patch.dict(os.environ, {"REDINK_TEST_TLS_PASSWORD": password}):
            port = self._start_https(replace(self.settings, tls_key_password_env="REDINK_TEST_TLS_PASSWORD"))
            context = ssl.create_default_context(cafile=self.ca_path)
            with urlopen(f"https://127.0.0.1:{port}/healthz", context=context, timeout=5) as response:
                self.assertEqual(json.load(response)["scheme"], "https")


if __name__ == "__main__":
    unittest.main()
