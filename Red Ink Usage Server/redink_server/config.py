"""Validated local configuration; credentials are separate from served INI files."""
from __future__ import annotations

import hashlib
import json
import os
import re
import secrets
import sys
import tomllib
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

from cryptography.fernet import Fernet


DEFAULT_TOML = '''# Red Ink Usage Server - restart after manual edits; safe UI updates apply immediately.
# Relative paths below are relative to this data directory.
[server]
host = "127.0.0.1"
port = 8765
threads = 8
# Add the hostname used by browsers/add-ins before exposing the server.
trusted_hosts = ["localhost", "127.0.0.1", "[::1]"]
# Set true when browsers connect through HTTPS (recommended for production).
secure_cookies = false
# Direct plain HTTP on a non-loopback interface requires this explicit opt-in.
allow_insecure_remote = false

[tls]
# Localhost HTTP needs no certificate. Set true for native HTTPS.
enabled = false
# PEM leaf certificate followed by intermediate certificates, and its PEM key.
# Relative paths are below this data directory. Keep them outside config-files.
certificate = ""
private_key = ""
# Optional environment variable holding an encrypted PEM key's password.
# Native HTTPS always enables secure cookies, regardless of server.secure_cookies.
key_password_env = ""

[collection]
# These two options are retained for v1.0 configuration compatibility.
# They are no longer used for daily or monthly counting.
recent_report_window_seconds = 43200
maximum_clock_skew_seconds = 120
maximum_request_bytes = 65536

[imports]
# Existing LicenseCounterPath shares and SharedLogger LogPath folders.
# A source must be readable by the account running the server.
counter_paths = []
# Directories/files containing existing redink-<hash>-WD/XL/OL.log files.
# INI_LogPath is a client setting; point these server paths to a readable share.
log_paths = []
recursive = true
scan_seconds = 60
# Counter-file size cap; for SharedLogger .log this is the per-scan read budget.
# Larger growing logs continue on subsequent scans from their last full line.
maximum_file_bytes = 16777216
maximum_files_per_scan = 10000

[publishing]
# Optional alternate directory for redink.ini and related files and signatures.
# Empty means the default config-files directory in the data directory.
source_directory = ""

[authentication]
# Protected ciphertext, changed exclusively from the authenticated Settings page.
passwordless_mode = ""

[retention]
# Earliest client-local date that the collector retains (YYYY-MM-DD; empty disables).
# This is a HARD ingestion floor: older offline replays are acknowledged and discarded.
cutoff_date = ""
# Optional rolling window in calendar months (0 disables). The newer floor wins.
keep_months = 0
# Optional scheduled cleanup on scans (false = manual cleanup from Settings).
auto_cleanup = false

[billing]
# Leave unit_price empty for a usage-only report. No tax is calculated.
unit_price = ""
currency = "CHF"
'''


class ConfigError(ValueError):
    pass


def default_data_dir() -> Path:
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
    return base / "data"


def token_hash(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def private_write(path: Path, text: str) -> None:
    """Write a new file exclusively; restrictive POSIX mode, Windows ACL inherited."""
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as output:
        output.write(text)
        output.flush()
        os.fsync(output.fileno())


def initialize(data_dir: Path) -> dict:
    data_dir = data_dir.resolve()
    data_dir.mkdir(parents=True, exist_ok=True)
    if any((data_dir / name).exists() for name in ("server.toml", "secrets.json", "usage.sqlite3")):
        raise ConfigError("This data directory is already initialized or contains a partial initialization. No files were overwritten.")
    from werkzeug.security import generate_password_hash
    password = secrets.token_urlsafe(24)
    ingestion_token = secrets.token_urlsafe(32)
    config_token = secrets.token_urlsafe(32)
    encryption_key = Fernet.generate_key().decode("ascii")
    credentials = {
        "session_secret": secrets.token_hex(32),
        "admin_password_hash": generate_password_hash(password),
        "ingestion_token_hash": token_hash(ingestion_token),
        "config_token_hash": token_hash(config_token),
        "configuration_encryption_key": encryption_key,
        "ingestion_token_ciphertext": Fernet(encryption_key.encode()).encrypt(ingestion_token.encode()).decode(),
        "config_token_ciphertext": Fernet(encryption_key.encode()).encrypt(config_token.encode()).decode(),
    }
    private_write(data_dir / "secrets.json", json.dumps(credentials, indent=2) + "\n")
    private_write(data_dir / "server.toml", DEFAULT_TOML)
    (data_dir / "config-files").mkdir(exist_ok=True)
    (data_dir / "backups").mkdir(exist_ok=True)
    return {"admin_password": password, "ingestion_token": ingestion_token, "config_token": config_token}


@dataclass(frozen=True)
class Settings:
    data_dir: Path
    credentials: dict = field(repr=False)
    host: str = "127.0.0.1"
    port: int = 8765
    threads: int = 8
    trusted_hosts: tuple[str, ...] = ("localhost", "127.0.0.1", "[::1]")
    secure_cookies: bool = False
    allow_insecure_remote: bool = False
    tls_enabled: bool = False
    tls_certificate: Path | None = None
    tls_private_key: Path | None = None
    tls_key_password_env: str = ""
    recent_report_window_seconds: int = 43200
    maximum_clock_skew_seconds: int = 120
    maximum_request_bytes: int = 65536
    counter_paths: tuple[Path, ...] = ()
    log_paths: tuple[Path, ...] = ()
    recursive: bool = True
    scan_seconds: int = 60
    maximum_file_bytes: int = 16777216
    maximum_files_per_scan: int = 10000
    publishing_source_directory: Path | None = None
    passwordless_mode: str = ""
    unit_price: str = ""
    currency: str = "CHF"
    cutoff_date: str = ""
    keep_months: int = 0
    auto_cleanup: bool = False

    @property
    def database_path(self) -> Path:
        return self.data_dir / "usage.sqlite3"

    @property
    def config_directory(self) -> Path:
        return self.publishing_source_directory or (self.data_dir / "config-files")


_KNOWN = {
    "server": {"host", "port", "threads", "trusted_hosts", "secure_cookies", "allow_insecure_remote"},
    "tls": {"enabled", "certificate", "private_key", "key_password_env"},
    "collection": {"recent_report_window_seconds", "maximum_clock_skew_seconds", "maximum_request_bytes"},
    "imports": {"counter_paths", "log_paths", "recursive", "scan_seconds", "maximum_file_bytes", "maximum_files_per_scan"},
    "billing": {"unit_price", "currency"},
    "publishing": {"source_directory"},
    "authentication": {"passwordless_mode"},
    "retention": {"cutoff_date", "keep_months", "auto_cleanup"},
}


def load_settings(data_dir: Path) -> Settings:
    data_dir = data_dir.resolve()
    try:
        with (data_dir / "server.toml").open("rb") as source:
            config = tomllib.load(source)
        credentials = json.loads((data_dir / "secrets.json").read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ConfigError("Cannot read server.toml and secrets.json. Run init for a new directory, or check the existing files.") from exc
    # Add a separate encryption key to pre-1.3 installs without touching any
    # password or collector-token hash. Existing raw tokens cannot be recovered.
    if isinstance(credentials, dict) and "configuration_encryption_key" not in credentials:
        upgraded = dict(credentials)
        upgraded["configuration_encryption_key"] = Fernet.generate_key().decode("ascii")
        temp = data_dir / (".secrets-upgrade-" + secrets.token_hex(8) + ".tmp")
        private_write(temp, json.dumps(upgraded, indent=2) + "\n")
        try:
            os.replace(temp, data_dir / "secrets.json")
        finally:
            temp.unlink(missing_ok=True)
        credentials = upgraded
    if not isinstance(credentials, dict) or any(not isinstance(credentials.get(k), str) or len(credentials[k]) < 32
        for k in ("session_secret", "admin_password_hash", "ingestion_token_hash", "config_token_hash")):
        raise ConfigError("secrets.json is incomplete or malformed.")
    values = {}
    for section, entries in config.items():
        if section not in _KNOWN or not isinstance(entries, dict):
            raise ConfigError(f"Unrecognized configuration section: {section}")
        for key, value in entries.items():
            if key not in _KNOWN[section]:
                raise ConfigError(f"Unrecognized configuration option: {section}.{key}")
            values[("tls_" + key if section == "tls" else "publishing_" + key if section == "publishing" else key)] = value
    ranges = {"port": (1, 65535), "threads": (1, 64), "recent_report_window_seconds": (60, 1814400),
              "maximum_clock_skew_seconds": (0, 3600), "maximum_request_bytes": (4096, 1048576),
              "scan_seconds": (5, 86400), "maximum_file_bytes": (1024, 1073741824),
              "maximum_files_per_scan": (1, 1000000), "keep_months": (0, 1200)}
    for key, (low, high) in ranges.items():
        if key in values and (type(values[key]) is not int or not low <= values[key] <= high):
            raise ConfigError(f"{key} must be an integer between {low} and {high}.")
    for key in ("secure_cookies", "allow_insecure_remote", "recursive", "tls_enabled", "auto_cleanup"):
        if key in values and type(values[key]) is not bool:
            raise ConfigError(f"{key} must be true or false.")
    from .retention import validate_cutoff_date
    if "cutoff_date" in values:
        values["cutoff_date"] = validate_cutoff_date(values["cutoff_date"])
    publishing = values.get("publishing_source_directory", "")
    if not isinstance(publishing, str) or "\x00" in publishing or len(publishing) > 4096:
        raise ConfigError("publishing.source_directory must be a valid directory path string.")
    values["publishing_source_directory"] = ((Path(publishing) if Path(publishing).is_absolute() else data_dir / publishing).resolve() if publishing.strip() else None)
    policy = values.get("passwordless_mode", "")
    if not isinstance(policy, str) or len(policy) > 1024:
        raise ConfigError("authentication.passwordless_mode must be a protected value.")
    if policy:
        from .settings_ui import decrypt_secret
        if decrypt_secret(credentials, policy) != "localhost-only-v1":
            raise ConfigError("Unrecognized protected sign-in policy.")
    for key in ("tls_certificate", "tls_private_key"):
        value = values.get(key, "")
        if not isinstance(value, str) or "\x00" in value:
            raise ConfigError(f"{key} must be a file path string.")
        values[key] = (Path(value) if Path(value).is_absolute() else data_dir / value) if value.strip() else None
    password_env = values.get("tls_key_password_env", "")
    if not isinstance(password_env, str) or (password_env and not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,127}", password_env)):
        raise ConfigError("tls.key_password_env must be an environment variable name, not a password.")
    if values.get("tls_enabled", False):
        if values["tls_certificate"] is None or values["tls_private_key"] is None:
            raise ConfigError("Native HTTPS requires both tls.certificate and tls.private_key.")
        values["secure_cookies"] = True
    for key in ("counter_paths", "log_paths"):
        paths = values.get(key, [])
        if not isinstance(paths, list) or len(paths) > 100 or any(not isinstance(p, str) or not p.strip() for p in paths):
            raise ConfigError(f"{key} must be a list of up to 100 nonempty paths.")
        values[key] = tuple(Path(p) if Path(p).is_absolute() else data_dir / p for p in paths)
    hosts = values.get("trusted_hosts", ["localhost", "127.0.0.1", "[::1]"])
    if not isinstance(hosts, list) or not hosts or any(not isinstance(h, str) or not h.strip() or "/" in h for h in hosts):
        raise ConfigError("trusted_hosts must contain explicit host names.")
    values["trusted_hosts"] = tuple(hosts)
    host = values.get("host", "127.0.0.1")
    if not isinstance(host, str) or not host.strip() or len(host) > 253:
        raise ConfigError("host must be a host name or IP address.")
    price = values.get("unit_price", "")
    if not isinstance(price, str):
        raise ConfigError("unit_price must be quoted decimal text or an empty string.")
    if price:
        try:
            amount = Decimal(price)
            if not amount.is_finite() or amount < 0 or amount > 1000000000 or amount.as_tuple().exponent < -6:
                raise InvalidOperation
        except InvalidOperation:
            raise ConfigError("unit_price must be a nonnegative finite decimal with at most six fractional digits.") from None
    currency = values.get("currency", "CHF")
    if not isinstance(currency, str) or len(currency) != 3 or not currency.isascii() or not currency.isalpha():
        raise ConfigError("currency must be a three-letter currency label.")
    values["currency"] = currency.upper()
    return Settings(data_dir=data_dir, credentials=credentials, **values)
