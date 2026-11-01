"""Authenticated English dashboard and backward-compatible Red Ink collector."""
from __future__ import annotations

import functools
import hmac
import json
import logging
import os
import re
import secrets
import sqlite3
import threading
import time
from collections import defaultdict, deque
from datetime import datetime, timedelta, timezone
from pathlib import Path

from flask import Flask, Response, abort, jsonify, redirect, render_template, request, send_file, session, url_for
from werkzeug.exceptions import BadRequest, HTTPException
from werkzeug.security import check_password_hash

from . import __version__
from .config import Settings, token_hash, private_write, ConfigError, load_settings
from .settings_ui import (configuration_document, update_configuration, update_passwordless,
                          is_passwordless, is_loopback_request, encrypt_secret, decrypt_secret, apply_live_settings, restore_configuration)
from .user_analytics import function_users, license_users

from .protocol import ProtocolError, normalize_event, parse_legacy_path
from .reports import daily_csv, monthly_report, range_months, report_csv, users_csv, valid_day, valid_month
from .store import Store, iso


CONFIG_FILES = ("redink.ini", "allmodels.ini", "specialservices.ini")
CONFIG_ALLOWLIST = CONFIG_FILES + tuple(name + ".sig" for name in CONFIG_FILES)


class LoginLimiter:
    def __init__(self):
        self.lock = threading.Lock()
        self.failures = defaultdict(deque)

    def blocked(self, address):
        now = time.monotonic()
        with self.lock:
            stale = [key for key, values in self.failures.items() if not values or values[-1] < now - 300]
            for key in stale:
                del self.failures[key]
            if len(self.failures) > 4096 and address not in self.failures:
                return True
            values = self.failures[address]
            while values and values[0] < now - 300:
                values.popleft()
            return len(values) >= 5

    def failed(self, address):
        with self.lock:
            self.failures[address].append(time.monotonic())


def _json_pairs(pairs):
    value = {}
    for key, item in pairs:
        if key in value:
            raise ProtocolError("Duplicate JSON field.")
        value[key] = item
    return value


def create_app(settings: Settings, store: Store | None = None) -> Flask:
    app = Flask(__name__, template_folder="templates", static_folder="static")
    app.config.update(
        SECRET_KEY=settings.credentials["session_secret"], MAX_CONTENT_LENGTH=settings.maximum_request_bytes,
        MAX_FORM_MEMORY_SIZE=settings.maximum_request_bytes, MAX_FORM_PARTS=80,
        SESSION_COOKIE_NAME="redink_usage_admin", SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SECURE=settings.secure_cookies, SESSION_COOKIE_SAMESITE="Strict",
        PERMANENT_SESSION_LIFETIME=timedelta(hours=8), TRUSTED_HOSTS=list(settings.trusted_hosts),
    )
    store = store or Store(settings)
    app.extensions["redink_store"] = store
    limiter = LoginLimiter()
    auth_policy = {"passwordless": is_passwordless(settings)}

    def local_passwordless_allowed():
        return (auth_policy["passwordless"] and
                settings.host in ("localhost", "127.0.0.1", "::1") and
                is_loopback_request(request.remote_addr) and
                request.host.split(":", 1)[0].lower() in ("localhost", "127.0.0.1"))


    def csrf_token():
        if "csrf" not in session:
            session["csrf"] = secrets.token_urlsafe(32)
        return session["csrf"]

    app.jinja_env.globals.update(csrf_token=csrf_token, app_version=__version__)

    def check_csrf():
        supplied = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token", "")
        if not supplied or not supplied.isascii() or not hmac.compare_digest(supplied, session.get("csrf", "")):
            abort(400, description="This form expired. Reload the page and try again.")

    def admin_required(function):
        @functools.wraps(function)
        def wrapped(*args, **kwargs):
            if not session.get("admin"):
                if request.path.startswith("/api/"):
                    return jsonify(error="Authentication required."), 401
                return redirect(url_for("login"))
            return function(*args, **kwargs)
        return wrapped

    def authenticated_token(supplied: str, kind: str) -> bool:
        if not supplied:
            bearer = request.headers.get("Authorization", "")
            supplied = bearer[7:] if bearer.startswith("Bearer ") else ""
        return 16 <= len(supplied) <= 256 and hmac.compare_digest(token_hash(supplied), settings.credentials[kind + "_token_hash"])

    @app.after_request
    def security_headers(response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Content-Security-Policy"] = "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
        response.headers["Cache-Control"] = "no-store"
        response.headers["X-Robots-Tag"] = "noindex, nofollow, noarchive"
        if settings.secure_cookies:
            response.headers["Strict-Transport-Security"] = "max-age=31536000"
        return response

    @app.errorhandler(HTTPException)
    def http_error(error):
        if request.path.startswith(("/api/", "/collect/", "/config/")):
            return jsonify(accepted=False, error=error.description), error.code
        return Response(f"Red Ink Usage Server: {error.name}. Reload the page or check the server configuration.", status=error.code, mimetype="text/plain")

    @app.errorhandler(sqlite3.Error)
    def database_error(error):
        app.logger.error("Database operation failed (%s); request was not acknowledged.", type(error).__name__)
        return jsonify(accepted=False, error="Storage is temporarily unavailable. Retry later."), 503

    @app.errorhandler(Exception)
    def unexpected_error(error):
        # Flask's default exception logging includes the request path. Collector
        # URLs contain credentials, so log only the exception class here.
        app.logger.error("Request failed (%s); no success acknowledgement was issued.", type(error).__name__)
        return jsonify(accepted=False, error="The request could not be completed. Review the server log and retry."), 500

    @app.route("/healthz")
    def healthz():
        # No paths, database contents, identities or configuration in public health.
        return jsonify(status="ok", version=__version__)

    @app.route("/login", methods=["GET", "POST"])
    def login():
        error = None
        status = 200
        if request.method == "POST":
            check_csrf()
            address = request.remote_addr or "unknown"
            if limiter.blocked(address):
                error, status = "Too many failed attempts. Try again in five minutes.", 429
            elif ((request.form.get("sign_in_mode") == "localhost" and local_passwordless_allowed()) or
                  (request.form.get("sign_in_mode") != "localhost" and
                   check_password_hash(settings.credentials["admin_password_hash"], request.form.get("password", "")[:1024]))):
                session.clear()
                session["admin"] = True
                session.permanent = True
                return redirect(url_for("index"))
            else:
                limiter.failed(address)
                error, status = "The password was not accepted.", 401
        return render_template("login.html", error=error, version=__version__, passwordless=local_passwordless_allowed()), status

    @app.post("/logout")
    @admin_required
    def logout():
        check_csrf()
        session.clear()
        return redirect(url_for("login"))

    @app.get("/")
    @admin_required
    def index():
        return render_template("index.html", version=__version__, current_month=store.clock().strftime("%Y-%m"))

    def config_metadata():
        rows = []
        root = settings.config_directory.resolve()
        for name in CONFIG_FILES:
            path = root / name
            available = path.is_file() and not path.is_symlink() and path.resolve().parent == root
            stat = path.stat() if available else None
            rows.append({"name": name, "available": available, "bytes": stat.st_size if stat else 0,
                         "modified_utc": iso(datetime.fromtimestamp(stat.st_mtime, timezone.utc)) if stat else None,
                         "url": "/admin/config/" + name if available else ""})
        return rows

    @app.get("/api/settings")
    @admin_required
    def read_settings_api():
        result = configuration_document(settings)
        result['passwordless_local_only'] = auth_policy['passwordless']
        result['default_config_directory'] = str(settings.data_dir / 'config-files')
        result['tokens'] = {}
        for kind in ('ingestion', 'config'):
            cipher = settings.credentials.get(kind + '_token_ciphertext', '')
            result['tokens'][kind] = decrypt_secret(settings.credentials, cipher) if cipher else None
        return jsonify(result)

    runtime_settings_lock = threading.RLock()

    def apply_runtime_settings(next_settings, *, scan_now=False):
        nonlocal settings
        with runtime_settings_lock:
            settings = next_settings
            store.settings = next_settings
            scanner = app.extensions.get('redink_scanner')
            if scanner is not None:
                scanner.update_settings(next_settings, scan_now=scan_now)
            app.config.update(
                MAX_CONTENT_LENGTH=next_settings.maximum_request_bytes,
                MAX_FORM_MEMORY_SIZE=next_settings.maximum_request_bytes,
                SESSION_COOKIE_SECURE=next_settings.secure_cookies,
                TRUSTED_HOSTS=list(next_settings.trusted_hosts),
            )
            auth_policy['passwordless'] = is_passwordless(next_settings)

    app.extensions['redink_apply_settings'] = apply_runtime_settings

    @app.get('/api/settings/listener')
    @admin_required
    def listener_state_api():
        supervisor = app.extensions.get('redink_listener')
        return jsonify(supervisor.status() if supervisor is not None else
                       {'status': 'unmanaged', 'message': 'No production listener supervisor is active.'})

    @app.post('/api/settings')
    @admin_required
    def save_settings_api():
        check_csrf()
        try:
            data = request.get_json(silent=False)
            if not isinstance(data, dict):
                raise ConfigError('Expected a JSON object.')
            with runtime_settings_lock:
                supervisor = app.extensions.get('redink_listener')
                if supervisor is not None and supervisor.is_pending():
                    return jsonify(error='A listener update is in progress. Wait for its result before saving again.'), 409
                previous_toml = (settings.data_dir / 'server.toml').read_text(encoding='utf-8')
                result = update_configuration(settings, data)
                try:
                    fresh = load_settings(settings.data_dir)
                    if fresh.host not in ('127.0.0.1', 'localhost', '::1') and not fresh.tls_enabled and not fresh.allow_insecure_remote:
                        raise ConfigError('A remote listener requires HTTPS or explicit allow_insecure_remote=true.')
                    if fresh.tls_enabled:
                        from .tls import validate_tls
                        validate_tls(fresh)
                except Exception:
                    restore_configuration(settings.data_dir, previous_toml)
                    raise
                # Scanner/publishing/billing/retention are live-safe. Security,
                # binding and TLS changes are committed by the supervisor only
                # after successful listener preparation; failure rolls back.
                live_settings, changes = apply_live_settings(settings, fresh, data)
                network_changed = any(getattr(fresh, field) != getattr(settings, field) for field in (
                    'host', 'port', 'threads', 'trusted_hosts', 'secure_cookies',
                    'allow_insecure_remote', 'tls_enabled', 'tls_certificate',
                    'tls_private_key', 'tls_key_password_env', 'maximum_request_bytes'))
                if network_changed and supervisor is not None:
                    if not supervisor.reconfigure(fresh, previous_toml):
                        restore_configuration(settings.data_dir, previous_toml)
                        return jsonify(error='Listener update could not be scheduled; no configuration was changed.'), 409
                    # Safe settings are also applied after a successful rebind.
                    applied_now = []
                    listener_update = 'pending'
                    restart_required = False
                elif network_changed:
                    # Flask's test client has no active listening socket. Do not
                    # claim the host/port has actually changed in that mode.
                    apply_runtime_settings(live_settings, scan_now='imports' in data)
                    applied_now = changes
                    listener_update = 'unmanaged'
                    restart_required = True
                else:
                    apply_runtime_settings(fresh, scan_now='imports' in data)
                    applied_now = sorted(set(changes) | set(data.get('collection', {})))
                    listener_update = 'not_needed'
                    restart_required = False
                active_for_warnings = fresh
                warnings = []
                if 'imports' in data:
                    for name, roots in (('LogPath', active_for_warnings.log_paths), ('LicenseCounter', active_for_warnings.counter_paths)):
                        for root in roots:
                            if not root.exists():
                                warnings.append(f'{name} source is not accessible: {root}. The scanner will retry when available.')
                if 'publishing' in data:
                    folder = active_for_warnings.config_directory
                    if not folder.is_dir():
                        warnings.append(f'Publishing folder does not exist: {folder}. Create it and place the INI files there.')
                    else:
                        missing = [name for name in CONFIG_FILES if not (folder / name).is_file()]
                        if missing:
                            warnings.append('Publishing folder has no ' + ', '.join(missing) + ' yet.')
                result.update(restart_required=restart_required, listener_update=listener_update,
                              applied_now=applied_now, warnings=warnings,
                              scan_scheduled=bool('imports' in data and not network_changed and app.extensions.get('redink_scanner') is not None),
                              published_directory=str(active_for_warnings.config_directory))
                return jsonify(result)
        except (ConfigError, ValueError, OSError) as exc:
            return jsonify(error=str(exc)), 400

    @app.post('/api/settings/passwordless')
    @admin_required
    def set_passwordless_api():
        check_csrf()
        try:
            if settings.host not in ('127.0.0.1', 'localhost', '::1'):
                raise ConfigError('Passwordless administration is only supported on a loopback listener.')
            if not is_loopback_request(request.remote_addr):
                raise ConfigError('Passwordless policy can only be changed locally.')
            requested = request.get_json(silent=False)
            if not isinstance(requested, dict) or type(requested.get('enabled')) is not bool:
                raise ConfigError('Expected a boolean enabled value.')
            updated = update_passwordless(settings, requested['enabled'])
            auth_policy['passwordless'] = is_passwordless(updated)
            return jsonify(enabled=auth_policy['passwordless'], saved=True)
        except (ConfigError, ValueError, OSError) as exc:
            return jsonify(error=str(exc)), 400

    @app.get('/api/settings/autostart')
    @admin_required
    def autostart_status():
        if not is_loopback_request(request.remote_addr):
            return jsonify(error='Autostart can only be managed locally on the server.'), 403
        from .startup import manage_autostart, StartupError
        try:
            return jsonify(manage_autostart(settings.data_dir, 'status'))
        except StartupError as exc:
            return jsonify(error=str(exc)), 409

    @app.post('/api/settings/autostart')
    @admin_required
    def autostart_change():
        check_csrf()
        if not is_loopback_request(request.remote_addr):
            return jsonify(error='Autostart can only be managed locally on the server.'), 403
        body = request.get_json(silent=True)
        if not isinstance(body, dict) or type(body.get('enabled')) is not bool:
            return jsonify(error='Expected a boolean enabled value.'), 400
        from .startup import manage_autostart, StartupError
        try:
            return jsonify(manage_autostart(settings.data_dir, 'enable' if body['enabled'] else 'disable'))
        except StartupError as exc:
            return jsonify(error=str(exc)), 409

    @app.post('/api/settings/shutdown')
    @admin_required
    def server_shutdown():
        check_csrf()
        if not is_loopback_request(request.remote_addr):
            return jsonify(error='Shut down is restricted to localhost administration.'), 403
        stop = app.extensions.get('redink_shutdown')
        if stop is None:
            return jsonify(error='Shut down is only available while running the production server.'), 409
        if not stop():
            return jsonify(error='Shut down has already been requested.'), 409
        return jsonify(accepted=True, message='Server shutdown requested. The dashboard will become unavailable.')

    @app.post('/api/settings/token')
    @admin_required
    def rotate_settings_token():
        check_csrf()
        data = request.get_json(silent=True) or {}
        kind = data.get('kind')
        if kind not in ('ingestion', 'config'):
            return jsonify(error='Invalid token type.'), 400
        from .settings_ui import rotate_token
        try:
            token = rotate_token(settings, kind)
        except (ConfigError, OSError) as exc:
            return jsonify(error=str(exc)), 400
        return jsonify(rotated=True, kind=kind, token=token,
                       message='Token changed immediately. Update every affected client or deployment URL. The previous token is no longer valid.')

    @app.post('/api/settings/open-folder')
    @admin_required
    def open_config_folder():
        check_csrf()
        data = request.get_json(silent=True) or {}
        if data.get('which') not in ('default', 'published'):
            return jsonify(error='Invalid directory selection.'), 400
        if not is_loopback_request(request.remote_addr) or os.name != 'nt':
            return jsonify(error='Opening a local server folder is supported on Windows localhost only.'), 403
        folder = settings.data_dir / 'config-files' if data['which'] == 'default' else settings.config_directory
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder.resolve()))
        return jsonify(opened=True)

    @app.get('/api/users/functions')
    @admin_required
    def function_user_api():
        try:
            return jsonify(function_users(store, request.args.get('period','30d'),
                                          request.args.get('host','ALL'), request.args.get('user') or None))
        except ValueError as exc:
            return jsonify(error=str(exc)), 400

    @app.get('/api/users/functions.csv')
    @admin_required
    def function_users_csv_api():
        from .user_analytics import user_summary_csv
        try:
            data = function_users(store, request.args.get('period','30d'), request.args.get('host','ALL'), all_rows=True)
            return Response(user_summary_csv(data), mimetype='text/csv',
                            headers={'Content-Disposition':'attachment; filename="redink-function-users.csv"'})
        except ValueError as exc:
            return jsonify(error=str(exc)), 400

    @app.get('/api/users/licenses')
    @admin_required
    def license_users_api():
        try:
            return jsonify(license_users(store, request.args.get('month', store.clock().strftime('%Y-%m'))))
        except ValueError as exc:
            return jsonify(error=str(exc)), 400

    @app.get('/api/users/licenses.csv')
    @admin_required
    def license_users_csv_api():
        from .user_analytics import license_users_csv
        try:
            return Response(license_users_csv(license_users(store, request.args.get('month', store.clock().strftime('%Y-%m')))),
                            mimetype='text/csv',headers={'Content-Disposition':'attachment; filename="redink-license-users.csv"'})
        except ValueError as exc:
            return jsonify(error=str(exc)), 400

    @app.get("/api/settings/retention/preview")
    @admin_required
    def retention_preview_api():
        return jsonify(store.retention_preview())

    @app.post("/api/settings/retention/clean")
    @admin_required
    def retention_clean_api():
        check_csrf()
        data = request.get_json(silent=True) or {}
        cutoff = store.retention_cutoff()
        if not cutoff or data.get("confirmation") != "DELETE BEFORE " + cutoff:
            return jsonify(error="Review the current cutoff and enter the exact confirmation text."), 400
        try:
            return jsonify(store.clean_retention(cutoff))
        except ValueError as exc:
            return jsonify(error=str(exc)), 409

    @app.get("/api/dashboard")
    @admin_required
    def dashboard():
        try:
            month = valid_month(request.args.get("month", store.clock().strftime("%Y-%m")))
            day = valid_day(request.args.get("day", store.clock().astimezone().date().isoformat()))
        except ValueError as exc:
            return jsonify(error=str(exc)), 400
        data = store.dashboard(month, day)
        data["config_files"] = config_metadata()
        data["warnings"] = [
            "Monthly figures are observed usage. Late offline reports can revise previous months; missing reports do not prove zero usage.",
            "Daily figures count explicit add-in load dates on each client's calendar. They do not measure current sessions or AI use; offline dates appear after synchronization.",
        ]
        if data["month"]["monthly_only_users"]:
            data["warnings"].append(f"{data['month']['monthly_only_users']} monthly identities have no explicit daily evidence. Legacy monthly reports still count for monthly license reporting.")
        data["warnings"] += [row["detail"] for row in data["sources"] if row.get("errors") and row.get("detail")]
        return jsonify(data)

    @app.get("/api/log-analysis")
    @admin_required
    def log_analysis():
        try:
            return jsonify(store.log_analysis(request.args.get("period", "30d"), request.args.get("host", "ALL")))
        except ValueError as exc:
            return jsonify(error=str(exc)), 400

    @app.get("/api/log-analysis.csv")
    @admin_required
    def log_analysis_csv():
        import csv
        import io
        try:
            data = store.log_analysis(request.args.get("period", "30d"), request.args.get("host", "ALL"))
        except ValueError as exc:
            return jsonify(error=str(exc)), 400
        out = io.StringIO(newline="")
        writer = csv.writer(out)
        writer.writerow(("Feature", "Category", "Is run start", "Observed activities", "Period", "Office host", "Date basis"))
        for item in data["features"]:
            writer.writerow((item["feature"],item["group_label"],"Yes" if item["agentic"] else "No",
                             item["count"],data["period"],data["host_filter"],"client-local"))
        return Response(out.getvalue(), mimetype="text/csv", headers={
            "Content-Disposition": "attachment; filename=redink-log-features.csv"})

    @app.route("/api/collect", methods=["GET", "POST"])
    @app.route("/collect/<token>", methods=["GET", "POST"])
    @app.route("/collect/<token>/<path:legacy_path>", methods=["GET"])
    def collect(token="", legacy_path=""):
        if request.method == "HEAD":
            abort(405)
        if not authenticated_token(token, "ingestion"):
            return jsonify(accepted=False, error="Invalid collector credential."), 401
        try:
            if legacy_path:
                event = parse_legacy_path("/" + legacy_path)
                if event is None:
                    raise ProtocolError("Unrecognized collector path.")
            else:
                if request.method == "GET":
                    raw = request.args.to_dict(flat=False)
                elif request.mimetype == "application/json":
                    try:
                        raw = json.loads(request.get_data(cache=False).decode("utf-8-sig"), object_pairs_hook=_json_pairs)
                    except (ValueError, UnicodeError, RecursionError) as exc:
                        raise ProtocolError(str(exc) if isinstance(exc, ProtocolError) else "Invalid UTF-8 JSON request.") from None
                elif request.mimetype == "application/x-www-form-urlencoded":
                    raw = request.form.to_dict(flat=False)
                else:
                    return jsonify(accepted=False, error="Use JSON or form-encoded content."), 415
                event = normalize_event(raw)
            result = store.ingest(event, live_request=True)
            return jsonify(result)
        except ProtocolError as exc:
            return jsonify(accepted=False, error=str(exc)), 400

    def serve_configuration(name):
        if name not in CONFIG_ALLOWLIST:
            abort(404)
        root = settings.config_directory.resolve()
        path = root / name
        if path.is_symlink() or not path.is_file() or path.resolve().parent != root:
            abort(404)
        if path.stat().st_size > 8 * 1024 * 1024:
            abort(413, description="Configuration file exceeds 8 MiB.")
        return send_file(path, mimetype="text/plain", as_attachment=False, download_name=name, conditional=True)

    @app.get("/config/<token>/<name>")
    def configuration(token, name):
        if not authenticated_token(token, "config"):
            abort(401, description="Invalid configuration credential.")
        return serve_configuration(name)

    @app.get("/api/config/<name>")
    def configuration_bearer(name):
        if not authenticated_token("", "config"):
            abort(401, description="Invalid configuration credential.")
        return serve_configuration(name)

    @app.get("/admin/config/<name>")
    @admin_required
    def admin_configuration(name):
        return serve_configuration(name)

    def selected_report():
        current = store.clock().strftime("%Y-%m")
        start = request.args.get("from", current)
        end = request.args.get("to", current)
        try:
            return monthly_report(store, start, end)
        except ValueError as exc:
            abort(400, description=str(exc))

    @app.get("/api/reports/monthly.csv")
    @admin_required
    def monthly_csv():
        report = selected_report()
        return Response(report_csv(report), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename=redink-usage-{report['from_month']}-{report['to_month']}.csv"})

    @app.get("/api/reports/monthly.json")
    @admin_required
    def monthly_json():
        return jsonify(selected_report())

    @app.get("/api/reports/users.csv")
    @admin_required
    def user_csv():
        try:
            month = valid_month(request.args.get("month", store.clock().strftime("%Y-%m")))
        except ValueError as exc:
            abort(400, description=str(exc))
        return Response(users_csv(store, month), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename=redink-users-{month}.csv"})

    @app.get("/api/reports/daily.csv")
    @admin_required
    def load_dates_csv():
        try:
            month = valid_month(request.args.get("month", store.clock().strftime("%Y-%m")))
        except ValueError as exc:
            abort(400, description=str(exc))
        return Response(daily_csv(store, month), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename=redink-loaded-days-{month}.csv"})

    @app.post("/api/reports/snapshot")
    @admin_required
    def snapshot():
        check_csrf()
        return jsonify(store.save_report_snapshot(selected_report())), 201

    @app.get("/reports/print")
    @admin_required
    def print_report():
        return render_template("report.html", report=selected_report(), version=__version__)

    return app
