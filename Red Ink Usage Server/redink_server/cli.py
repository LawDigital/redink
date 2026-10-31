"""Console entry point shared by Python and the PyInstaller executable."""
from __future__ import annotations

import argparse
import getpass
import json
import logging
import logging.handlers
import os
import re
import secrets
import sqlite3
import threading
import time
import sys
from dataclasses import replace
from pathlib import Path

from werkzeug.security import generate_password_hash

from . import __version__
from .browser_start import browser_requested, dashboard_url, open_after_bind
from .config import ConfigError, default_data_dir, initialize, load_settings, private_write, token_hash
from .reports import monthly_report, report_csv
from .scanner import Scanner
from .store import Store
from .tls import create_http_server, create_https_server, import_pfx, validate_tls
from .web import create_app


class InstanceLock:
    """OS lock, released after crashes as well as orderly shutdown (no stale PID gate)."""
    def __init__(self, path):
        self.path = path
        self.stream = None

    def __enter__(self):
        self.stream = self.path.open("a+b")
        self.stream.seek(0)
        if self.stream.read(1) == b"":
            self.stream.write(b"0")
            self.stream.flush()
        self.stream.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(self.stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self.stream.close()
            raise ConfigError("Another server is already using this data directory.") from None
        return self

    def __exit__(self, *_):
        if self.stream:
            self.stream.close()


def configure_logging(data_dir):
    handler = logging.handlers.RotatingFileHandler(data_dir / "server.log", maxBytes=2 * 1024 * 1024, backupCount=5, encoding="utf-8")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s",
                        handlers=[logging.StreamHandler(), handler], force=True)
    # Waitress has no access logger by default. Do not add one containing token URLs.


def _credentials_update(settings, mutate):
    values = dict(settings.credentials)
    mutate(values)
    temporary = settings.data_dir / (".secrets-" + secrets.token_hex(8) + ".tmp")
    private_write(temporary, json.dumps(values, indent=2) + "\n")
    try:
        os.replace(temporary, settings.data_dir / "secrets.json")
    finally:
        temporary.unlink(missing_ok=True)


def parser():
    root = argparse.ArgumentParser(description="Red Ink Usage Server - monthly usage, reports, configuration delivery and log summaries.")
    root.add_argument("--version", action="version", version=__version__)
    commands = root.add_subparsers(dest="command", required=True)
    for name, help_text in (("init", "Initialize a new data directory and print initial credentials."),
                            ("serve", "Run the HTTP or explicitly configured HTTPS server."),
                            ("archive-logs", "Preview or archive verified closed monthly LogPath segments."),
                            ("check", "Validate configuration and database integrity."),
                            ("import", "Import existing monthly/daily evidence or SharedLogger metadata, preserving original dates."),
                            ("backup", "Create and verify a consistent SQLite online backup."),
                            ("report", "Export a monthly usage report and retain an immutable snapshot."),
                            ("set-password", "Set the administrator password; restart the server afterwards."),
                            ("rotate-token", "Replace an ingestion or configuration credential; restart afterwards."),
                            ("import-pfx", "Import a PFX/P12 into a new local certificate folder with an encrypted PEM key.")):
        sub = commands.add_parser(name, help=help_text)
        sub.add_argument("--data-dir", type=Path, default=default_data_dir())
        if name == "serve":
            sub.add_argument("--browser", action=argparse.BooleanOptionalAction, default=None,
                             help="Open the dashboard after binding (default: interactive console only).")
            sub.add_argument("--host")
            sub.add_argument("--port", type=int)
        elif name == "import":
            sub.add_argument("--path", required=True, type=Path)
            sub.add_argument("--kind", required=True, choices=("counter", "log"))
            sub.add_argument("--recursive", action="store_true")
            sub.add_argument("--force", action="store_true")
        elif name == "archive-logs":
            boundary = sub.add_mutually_exclusive_group(required=True)
            boundary.add_argument("--before", help="Exclusive YYYY-MM boundary (not in the future).")
            boundary.add_argument("--keep-months", type=int, help="Keep this many recent calendar months including the current month.")
            sub.add_argument("--destination", type=Path, required=True)
            sub.add_argument("--apply", action="store_true", help="Verify a compressed copy and remove the closed original.")
        elif name == "backup":
            sub.add_argument("--output", required=True, type=Path)
        elif name == "report":
            sub.add_argument("--from", required=True, dest="from_month")
            sub.add_argument("--to", required=True, dest="to_month")
            sub.add_argument("--format", choices=("csv", "json"), default="csv")
            sub.add_argument("--output", required=True, type=Path)
        elif name == "rotate-token":
            sub.add_argument("--kind", required=True, choices=("ingestion", "config"))
        elif name == "import-pfx":
            sub.add_argument("--source", required=True, type=Path, help="PFX/P12 certificate file containing its private key.")
            sub.add_argument("--name", default="server", help="New certificate folder name; an existing folder is never replaced.")
            sub.add_argument("--pfx-password-env", default="", help="Name of an environment variable containing the PFX password; otherwise prompt.")
            sub.add_argument("--key-password-env", default="REDINK_TLS_KEY_PASSWORD", help="Variable for the encrypted exported key's password; prompt if not set.")
    return root


def main(argv=None):
    args = parser().parse_args(argv)
    try:
        if args.command == "init":
            credentials = initialize(args.data_dir)
            print("Red Ink Usage Server initialized. Store these credentials securely; they are shown only now.")
            print("Administrator password: " + credentials["admin_password"])
            print("Ingestion token: " + credentials["ingestion_token"])
            print("Configuration token: " + credentials["config_token"])
            print("\nLocal collector URL (no certificate needed): http://127.0.0.1:8765/collect/" + credentials["ingestion_token"])
            print("Local INI URL: http://127.0.0.1:8765/config/" + credentials["config_token"] + "/redink.ini")
            print("For other computers, use the server's network hostname. Configure [tls] for native HTTPS, or use an HTTPS reverse proxy.")
            print("Place approved INI files and their .sig files in: " + str(args.data_dir.resolve() / "config-files"))
            print("Start: RedInkUsageServer.exe serve --data-dir \"" + str(args.data_dir.resolve()) + "\"")
            return 0
        settings = load_settings(args.data_dir)
        if args.command == "import-pfx":
            for name in (args.pfx_password_env, args.key_password_env):
                if name and not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,127}", name):
                    raise ConfigError("Password environment-variable arguments must contain variable names, not passwords.")
            if not args.key_password_env:
                raise ConfigError("--key-password-env must name the variable that will unlock the exported PEM key.")
            if args.pfx_password_env:
                if args.pfx_password_env not in os.environ:
                    raise ConfigError("The named PFX-password environment variable is missing.")
                pfx_password = os.environ[args.pfx_password_env]
            elif sys.stdin.isatty():
                pfx_password = getpass.getpass("PFX/P12 password (press Enter if none): ")
            else:
                raise ConfigError("Non-interactive PFX import requires --pfx-password-env with a defined environment variable.")
            key_password = os.environ.get(args.key_password_env)
            if key_password is None:
                if not sys.stdin.isatty():
                    raise ConfigError("Non-interactive PFX import requires the named key-password environment variable.")
                key_password = getpass.getpass("New encrypted PEM key password (at least 12 characters): ")
                if key_password != getpass.getpass("Repeat PEM key password: "):
                    raise ConfigError("The PEM key passwords do not match.")
            installed = import_pfx(settings.data_dir, args.source, pfx_password, key_password, name=args.name)
            print("Certificate imported. The private key is encrypted; no server settings were changed.")
            print("Edit the [tls] section in server.toml:")
            print("enabled = true")
            print("certificate = " + json.dumps(installed["certificate_relative"]))
            print("private_key = " + json.dumps(installed["private_key_relative"]))
            print("key_password_env = " + json.dumps(args.key_password_env))
            print("Set that environment variable for the server account to the PEM password, run check, then restart the server.")
            return 0
        if args.command == "set-password":
            first = getpass.getpass("New administrator password (at least 16 characters): ")
            second = getpass.getpass("Repeat password: ")
            if first != second or len(first) < 16:
                raise ConfigError("Passwords must match and contain at least 16 characters.")
            _credentials_update(settings, lambda value: value.update(admin_password_hash=generate_password_hash(first), session_secret=secrets.token_hex(32)))
            print("Password updated. Restart the server; existing browser sessions will be invalidated.")
            return 0
        if args.command == "rotate-token":
            # Keep the encrypted display copy and the actual token hash in
            # sync, including when rotation is initiated outside the UI.
            from .settings_ui import rotate_token
            value = rotate_token(settings, args.kind)
            print("New " + args.kind + " token: " + value)
            print("Update clients and restart the server. The previous token will stop working after restart.")
            return 0
        # Configure the server log BEFORE opening the database: automatic
        # schema upgrades and their backup paths must be visible in the log.
        if args.command == "serve":
            configure_logging(settings.data_dir)
        store = Store(settings)
        if args.command == "check":
            tls_metadata = validate_tls(settings)
            if store.check() != "ok":
                raise ConfigError("Database integrity check did not return ok.")
            from .store import CURRENT_SCHEMA_VERSION
            print("Configuration valid. Database quick_check: ok. Version " + __version__ +
                  "; schema v" + str(CURRENT_SCHEMA_VERSION))
            if tls_metadata["enabled"]:
                print("Native HTTPS certificate and matching key: valid. Expires " + tls_metadata["not_valid_after_utc"])
                print("Certificate SAN names: " + ", ".join(tls_metadata["dns_names"] + tls_metadata["ip_addresses"]))
            if settings.host not in {"127.0.0.1", "localhost", "::1"} and not settings.tls_enabled and not settings.allow_insecure_remote:
                print("Network binding is blocked: enable native [tls], use a local HTTPS reverse proxy, or explicitly opt in to allow_insecure_remote on a trusted network.")
                return 1
            return 0
        if args.command == "archive-logs":
            from .archive_logs import archive_segments, month_boundary
            before = args.before or month_boundary(args.keep_months)
            result = archive_segments(store,before,args.destination,apply=args.apply)
            print(json.dumps(result,indent=2))
            return int(bool(result["held"]))
        if args.command == "backup":
            store.backup(args.output)
            print("Verified database backup written to " + str(args.output.resolve()))
            print("Back up server.toml, secrets.json and config-files separately with restricted access.")
            return 0
        if args.command == "import":
            result = Scanner(settings, store).scan_path(args.path, args.kind, recursive=args.recursive, force=args.force)
            print(json.dumps(result, indent=2))
            return 1 if result["errors"] else 0
        if args.command == "report":
            if args.output.exists():
                raise ConfigError("Report destination already exists. Choose a new filename.")
            report = monthly_report(store, args.from_month, args.to_month)
            snapshot = store.save_report_snapshot(report)
            args.output.parent.mkdir(parents=True, exist_ok=True)
            content = report_csv(report) if args.format == "csv" else json.dumps(snapshot, indent=2, ensure_ascii=False) + "\n"
            private_write(args.output, content)
            print("Report " + snapshot["report_id"] + " exported. Observed monthly units: " + str(report["units"]))
            print("Late offline reports may revise these months. Retained snapshot SHA-256: " + snapshot["sha256"])
            return 0
        if args.command == "serve":
            settings = replace(settings, host=args.host if args.host is not None else settings.host, port=args.port if args.port is not None else settings.port)
            if not 1 <= settings.port <= 65535:
                raise ConfigError("port must be between 1 and 65535.")
            if settings.host not in {"127.0.0.1", "localhost", "::1"} and not settings.tls_enabled and not settings.allow_insecure_remote:
                raise ConfigError("Remote plain HTTP binding requires explicit allow_insecure_remote=true. Enable native [tls] or use a local HTTPS reverse proxy for HTTPS.")
            from .listener_runtime import ListenerSupervisor
            from .settings_ui import restore_configuration
            with InstanceLock(settings.data_dir / "server.lock"):
                app = create_app(settings, store)
                scanner = Scanner(settings, store)
                app.extensions['redink_scanner'] = scanner

                def listener_factory(next_settings):
                    return (create_https_server(next_settings, app) if next_settings.tls_enabled
                            else create_http_server(next_settings, app))

                def apply_runtime(next_settings):
                    apply = app.extensions['redink_apply_settings']
                    apply(next_settings)

                def rollback(previous_settings, previous_toml):
                    restore_configuration(previous_settings.data_dir, previous_toml)
                    apply_runtime(previous_settings)

                supervisor = ListenerSupervisor(settings, listener_factory, apply_runtime, rollback)
                app.extensions['redink_listener'] = supervisor
                app.extensions['redink_shutdown'] = supervisor.shutdown
                scanner.start()
                try:
                    def on_ready(active_settings):
                        try:
                            url = dashboard_url(active_settings)
                            print(f"Red Ink Usage Server {__version__}: {url}", flush=True)
                            if browser_requested(args.browser):
                                open_after_bind(url)
                        except (OSError, RuntimeError, ValueError) as exc:
                            logging.getLogger(__name__).warning('Open the server page manually (%s).', type(exc).__name__)
                    supervisor.run(on_ready=on_ready)
                finally:
                    scanner.stop()
            return 0
    except sqlite3.Error:
        print("Error: Database operation failed. Check local disk space, permissions and database integrity; no success was reported.", file=sys.stderr)
        return 1
    except (ConfigError, ValueError, OSError, RuntimeError) as exc:
        print("Error: " + str(exc), file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("Server stopped.")
        return 0
    return 1
