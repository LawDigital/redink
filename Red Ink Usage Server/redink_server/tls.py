"""Native HTTPS and explicit PFX import, separate from published INI files.

HTTP and native HTTPS both use Cheroot's WSGI listener to support controlled
hot-rebinding without closing Waitress trigger sockets from worker threads. Credentials and certificates
are never sent to an external certificate-management service.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import ssl
import sys
import threading
import time
import errno
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from cryptography import x509
from cryptography.exceptions import UnsupportedAlgorithm
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.serialization import pkcs12

from .config import ConfigError, Settings, private_write

logger = logging.getLogger(__name__)
MAX_CERTIFICATE_BYTES = 8 * 1024 * 1024


def _certificate_file(settings: Settings, path: Path | None, label: str) -> Path:
    if path is None:
        raise ConfigError(f"Native HTTPS requires {label}.")
    resolved = path.resolve()
    if resolved.is_relative_to(settings.config_directory.resolve()):
        raise ConfigError("TLS certificates and private keys must be outside the published config-files directory.")
    try:
        if not resolved.is_file() or resolved.stat().st_size > MAX_CERTIFICATE_BYTES:
            raise ConfigError(f"The {label} must be a regular file no larger than 8 MiB.")
    except OSError:
        raise ConfigError(f"Cannot read the {label}; check its path and server-account permissions.") from None
    return resolved


def _key_password(settings: Settings) -> str | None:
    name = settings.tls_key_password_env
    if not name:
        return None
    value = os.environ.get(name)
    if not value:
        raise ConfigError("The environment variable named by tls.key_password_env is missing or empty.")
    return value


def _validate_certificate_validity(certificate: x509.Certificate, now: datetime | None = None) -> None:
    now = now or datetime.now(timezone.utc)
    if not certificate.not_valid_before_utc <= now < certificate.not_valid_after_utc:
        raise ConfigError("The TLS certificate is expired or not valid yet. Install a currently valid certificate and restart.")


def validate_tls(settings: Settings) -> dict:
    """Validate files, certificate dates and key match without binding a socket.

    Trust and hostname validation remain the connecting client's responsibility;
    a self-signed/internal-CA certificate is supported when clients trust it.
    """
    if not settings.tls_enabled:
        return {"enabled": False}
    certificate = _certificate_file(settings, settings.tls_certificate, "TLS certificate")
    private_key = _certificate_file(settings, settings.tls_private_key, "TLS private key")
    try:
        chain = x509.load_pem_x509_certificates(certificate.read_bytes())
        if not chain:
            raise ValueError("No certificate")
        for cert in chain:
            _validate_certificate_validity(cert)
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.minimum_version = ssl.TLSVersion.TLSv1_2
        # A callback returning an empty password prevents OpenSSL from prompting
        # unexpectedly if an encrypted key has no configured password variable.
        password = _key_password(settings)
        context.load_cert_chain(str(certificate), str(private_key), password=lambda: password or "")
        try:
            alternatives = chain[0].extensions.get_extension_for_class(x509.SubjectAlternativeName).value
            dns_names = alternatives.get_values_for_type(x509.DNSName)
            ip_addresses = [str(value) for value in alternatives.get_values_for_type(x509.IPAddress)]
        except x509.ExtensionNotFound:
            dns_names, ip_addresses = [], []
        return {"enabled": True, "certificate": str(certificate), "private_key": str(private_key),
                "not_valid_after_utc": chain[0].not_valid_after_utc.isoformat(),
                "dns_names": dns_names, "ip_addresses": ip_addresses}
    except ConfigError:
        raise
    except (OSError, ValueError, UnsupportedAlgorithm):
        raise ConfigError("Cannot load the TLS certificate/private-key pair. Check PEM format, matching key, and encrypted-key password. HTTPS was not started.") from None


class TransportDiagnostics:
    """Classify Cheroot events without logging request text, URLs or exception messages."""
    def __init__(self):
        self._lock = threading.Lock()
        self._buckets = {}

    def report(self, msg, level, with_traceback):
        exc_type, exc, tb = sys.exc_info()
        category = "transport-notice"
        if isinstance(msg, str) and msg.startswith('Client ') and 'attempted to speak plain HTTP' in msg:
            category = "http-on-https-port"
        elif isinstance(msg, str) and msg.startswith('Client ') and 'during handshake:' in msg:
            category = "tls-handshake-aborted"
        elif with_traceback or exc is not None or level >= logging.WARNING:
            category = "request-processing-error"
        elif msg in ('Keyboard Interrupt: shutting down', 'SystemExit raised: shutting down'):
            category = "listener-stopping"
        error_type = type(exc).__name__ if exc is not None else "none"
        cause = exc
        seen = set()
        for _ in range(4):
            if cause is None or id(cause) in seen:
                break
            seen.add(id(cause))
            nested = cause.__cause__ or cause.__context__
            if nested is None:
                break
            cause = nested
        cause_type = type(cause).__name__ if cause is not None else "none"
        facts = cause if category in ('tls-handshake-aborted', 'http-on-https-port') else exc
        error_number = getattr(facts, 'errno', None)
        error_number = error_number if isinstance(error_number, int) else None
        # OpenSSL's reason code is a symbolic identifier, never its exception text.
        reason = getattr(facts, 'reason', '')
        reason = reason if isinstance(reason, str) and re.fullmatch(r'[A-Z0-9_]{1,80}', reason) else "none"
        if isinstance(exc, (ConnectionResetError, ConnectionAbortedError, BrokenPipeError, TimeoutError)) or (
            isinstance(exc, OSError) and error_number in (errno.ECONNRESET, errno.ECONNABORTED, errno.EPIPE, errno.ETIMEDOUT)):
            category = "client-disconnected"
            level = logging.DEBUG
        elif category in ('tls-handshake-aborted', 'http-on-https-port'):
            level = logging.INFO
        frames = []
        if with_traceback:
            while tb is not None:
                code = tb.tb_frame.f_code
                # Code locations only. Do not include source lines, local values or paths.
                module = tb.tb_frame.f_globals.get('__name__', '')
                if isinstance(module, str) and re.fullmatch(r'[A-Za-z0-9_.]+', module):
                    frames.append(module + '.' + code.co_name + ':' + str(tb.tb_lineno))
                tb = tb.tb_next
        location = ' > '.join(frames[-6:]) or 'none'
        # Bound the key space and report one occurrence per category per minute.
        # Exception facts come only from runtime types/codes, never raw messages.
        key = (category, level, error_type, cause_type, error_number, reason, location)
        now = time.monotonic()
        with self._lock:
            previous, suppressed = self._buckets.get(key, (-float('inf'), 0))
            if now - previous < 60:
                self._buckets[key] = (previous, suppressed + 1)
                return
            if key not in self._buckets and len(self._buckets) >= 64:
                self._buckets.pop(next(iter(self._buckets)))
            self._buckets[key] = (now, 0)
        action = {
            'http-on-https-port': 'Use https:// for this listener; check client, proxy and health-check URLs.',
            'tls-handshake-aborted': 'Check client certificate trust, server hostname/chain and probe configuration.',
            'request-processing-error': 'Check the exception code and code locations; inspect the failing client or handler.',
            'client-disconnected': 'The client closed or timed out the connection.',
            'listener-stopping': 'The listener is stopping.',
            'transport-notice': 'The listener reported a transport notice.',
        }[category]
        logger.log(level, 'Transport category=%s exception=%s cause=%s errno=%s ssl_reason=%s locations=%s '
                   'suppressed_since_previous=%s. %s', category, error_type, cause_type, error_number,
                   reason, location, suppressed, action)


def _create_listener(settings: Settings, app):
    """Construct (but do not bind) the shared HTTP/HTTPS transport.

    A single listener implementation is important: the supervisor must be able
    to stop either protocol through the same well-defined server lifecycle.
    """
    from cheroot.wsgi import Server

    diagnostics = TransportDiagnostics()

    class PrivateLogServer(Server):
        def error_log(self, msg="", level=logging.INFO, traceback=False):
            # Keep bearer URLs and raw exception/trace text out of operational logs.
            diagnostics.report(msg, level, traceback)

    server = PrivateLogServer((settings.host, settings.port), app, numthreads=settings.threads,
                              timeout=30, shutdown_timeout=5, request_queue_size=128,
                              accepted_queue_size=256, accepted_queue_timeout=10)
    server.max_request_body_size = settings.maximum_request_bytes
    server.max_request_header_size = 16384
    server.software = "RedInkUsageServer"
    return server


def create_http_server(settings: Settings, app):
    """Return an unbound plain HTTP listener (only if explicitly permitted)."""
    if settings.tls_enabled:
        raise ConfigError("Plain HTTP cannot be used when native HTTPS is enabled.")
    if settings.host not in ("127.0.0.1", "localhost", "::1") and not settings.allow_insecure_remote:
        raise ConfigError("Remote HTTP requires explicit allow_insecure_remote=true or HTTPS.")
    return _create_listener(settings, app)


def create_https_server(settings: Settings, app):
    """Return an unbound native HTTPS listener. Never fall back to HTTP."""
    if not settings.tls_enabled:
        raise ConfigError("Native HTTPS is disabled in this configuration.")
    metadata = validate_tls(settings)
    from cheroot.ssl.builtin import BuiltinSSLAdapter
    server = _create_listener(settings, app)
    try:
        server.ssl_adapter = BuiltinSSLAdapter(metadata["certificate"], metadata["private_key"],
                                               private_key_password=lambda: _key_password(settings) or "")
        server.ssl_adapter.context.minimum_version = ssl.TLSVersion.TLSv1_2
        server.ssl_adapter.context.set_alpn_protocols(["http/1.1"])
    except (OSError, ValueError, UnsupportedAlgorithm):
        raise ConfigError("TLS initialization failed. No HTTP listener was substituted.") from None
    return server


def import_pfx(data_dir: Path, pfx_path: Path, pfx_password: str | None,
               key_password: str, *, name: str = "server") -> dict:
    """Create a new named certificate folder; never overwrite an existing pair.

    The exported key is always encrypted with the supplied password. Passwords
    are accepted in memory, never as command-line values or returned metadata.
    """
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}", name):
        raise ConfigError("Certificate name must start with a letter or number and contain only letters, numbers, underscores or hyphens.")
    if not isinstance(key_password, str) or len(key_password) < 12:
        raise ConfigError("The exported private key requires a password of at least 12 characters.")
    data_dir = data_dir.resolve()
    destination = data_dir / "tls" / name
    if destination.exists():
        raise ConfigError("This named certificate folder already exists. Use a new name for renewal; no files were overwritten.")
    try:
        if not pfx_path.is_file() or pfx_path.stat().st_size > MAX_CERTIFICATE_BYTES:
            raise ConfigError("The PFX/P12 source must be a regular file no larger than 8 MiB.")
        key, leaf, additional = pkcs12.load_key_and_certificates(
            pfx_path.read_bytes(), pfx_password.encode("utf-8") if pfx_password else None)
        if key is None or leaf is None:
            raise ConfigError("The PFX/P12 file must contain a certificate and its private key.")
        _validate_certificate_validity(leaf)
        if key.public_key().public_bytes(serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo) != leaf.public_key().public_bytes(serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo):
            raise ConfigError("The PFX private key does not match its certificate.")
        chain = [leaf] + list(additional or [])
        for cert in chain:
            _validate_certificate_validity(cert)
        certificate_bytes = b"".join(cert.public_bytes(serialization.Encoding.PEM) for cert in chain)
        key_bytes = key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8,
                                      serialization.BestAvailableEncryption(key_password.encode("utf-8")))
    except ConfigError:
        raise
    except (OSError, ValueError, TypeError, UnsupportedAlgorithm):
        raise ConfigError("Cannot import PFX/P12. Check the source file and its password; no certificate was installed.") from None
    destination.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=".import-", dir=destination.parent))
    try:
        private_write(staging / "fullchain.pem", certificate_bytes.decode("ascii"))
        private_write(staging / "private-key.pem", key_bytes.decode("ascii"))
        # Validate the pair before making the named folder available.
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(str(staging / "fullchain.pem"), str(staging / "private-key.pem"), password=key_password)
        if destination.exists():
            raise ConfigError("This named certificate folder was created concurrently. No files were overwritten.")
        staging.rename(destination)
        return {"certificate": str(destination / "fullchain.pem"), "private_key": str(destination / "private-key.pem"),
                "certificate_relative": "tls/" + name + "/fullchain.pem", "private_key_relative": "tls/" + name + "/private-key.pem"}
    finally:
        if staging.exists():
            shutil.rmtree(staging)
