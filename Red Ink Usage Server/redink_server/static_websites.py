"""Read-only static company websites, on dedicated or shared ports.

Only configured document roots are visible. Shared-port hosting uses strict
prefix dispatch and excludes administrator/collector routes. Separately bound
website ports never receive the administrator application.
"""
from __future__ import annotations

import logging
import threading
from dataclasses import replace
from pathlib import Path

from .config import ConfigError, Settings, StaticWebsite, RESERVED_SERVER_URL_ROOTS
from .tls import create_http_server, create_https_server

logger = logging.getLogger(__name__)


BLOCKED_EXTENSIONS = frozenset({
    '.key', '.pem', '.p12', '.pfx', '.cer', '.crt', '.der', '.jks', '.kdbx',
    '.sqlite', '.sqlite3', '.db', '.toml', '.ini', '.sig', '.env', '.py', '.pyc',
    '.ps1', '.psm1', '.bat', '.cmd', '.sh', '.vbs', '.config', '.bak', '.backup', '.log',
    '.php', '.phtml', '.jsp', '.asp', '.aspx', '.cgi', '.pl', '.rb', '.sql',
    '.cs', '.vb', '.vbproj', '.csproj', '.sln',
})


def _inside(path: Path, root: Path) -> bool:
    return path == root or root in path.parents



def resolve_static_file(site: StaticWebsite, relative: str, trailing_slash: bool):
    """Resolve only files below the approved root; return (path, redirect_slash).

    Returning None for a missing/denied file prevents sensitive-path details
    from being exposed in HTTP responses. No directory enumeration is provided.
    """
    if "\\" in relative or ":" in relative or "\x00" in relative or "//" in relative:
        return None, False
    pieces = [piece for piece in relative.split("/") if piece]
    if any(piece in (".", "..") or piece.startswith(".") for piece in pieces):
        return None, False
    if pieces and Path(pieces[-1]).suffix.lower() in BLOCKED_EXTENSIONS:
        return None, False
    source = site.directory.joinpath(*pieces)
    try:
        target = source.resolve(strict=True)
    except (OSError, RuntimeError):
        target = None
    if target is not None and not _inside(target, site.directory):
        return None, False
    if target is not None and target.is_dir():
        if not trailing_slash:
            return None, True
        target = next((candidate.resolve() for index in ("index.html", "index.htm")
                       if (candidate := target / index).is_file()), None)
    if (target is None or not target.is_file()) and site.spa_fallback and pieces and "." not in pieces[-1]:
        fallback = site.directory / "index.html"
        target = fallback.resolve() if fallback.is_file() else None
    if target is None or not _inside(target, site.directory) or not target.is_file():
        return None, False
    if target.suffix.lower() in BLOCKED_EXTENSIONS or any(p.startswith(".") for p in target.relative_to(site.directory).parts):
        return None, False
    return target, False


def create_static_application(sites: tuple[StaticWebsite, ...]):
    """Flask app with deterministic prefix routing, longest match first."""
    from flask import Flask, abort, redirect, request, send_file

    app = Flask(__name__, static_folder=None)
    prefixes = tuple(sorted(sites, key=lambda item: len(item.path), reverse=True))

    @app.after_request
    def add_headers(response):
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['Referrer-Policy'] = 'no-referrer'
        response.headers['X-Robots-Tag'] = 'noindex, nofollow'
        response.headers['X-Frame-Options'] = 'SAMEORIGIN'
        return response

    @app.route('/', defaults={'resource': ''}, methods=['GET'], provide_automatic_options=False)
    @app.route('/<path:resource>', methods=['GET'], provide_automatic_options=False)
    def serve_static(resource):
        # Flask's decoded path is used only after strict lexical checks. All
        # files are resolved again below, including symlink target checks.
        url = request.path
        for site in prefixes:
            prefix = site.path
            if prefix != '/' and url == prefix:
                return redirect(prefix + '/' + ('?' + request.query_string.decode('ascii', 'ignore') if request.query_string else ''), code=308)
            if prefix == '/' or url.startswith(prefix + '/'):
                relative = url[len(prefix):].lstrip('/') if prefix != '/' else url.lstrip('/')
                resolved, redirect_slash = resolve_static_file(site, relative, url.endswith('/'))
                if redirect_slash:
                    return redirect(url + '/', code=308)
                if resolved is None:
                    abort(404)
                response = send_file(resolved, conditional=True, etag=True)
                response.cache_control.no_cache = resolved.suffix.lower() in ('.html', '.htm')
                response.cache_control.max_age = 0 if resolved.suffix.lower() in ('.html', '.htm') else 3600
                return response
        abort(404)

    return app


class WebsiteListenerGroup:
    """One lifecycle shared by the collector and all additional HTTP servers.

    Every listener binds in prepare(), before any site becomes active. On
    conflict all prepared listeners are closed, so ListenerSupervisor can
    restore the previous generation and TOML atomically.
    """
    def __init__(self, listeners):
        self._listeners = list(listeners)
        self._stop = threading.Event()
        self._lock = threading.Lock()
        self._stopped = False
        self._threads = []
        self._fault = None
        self._prepared = []

    def prepare(self):
        try:
            for listener in self._listeners:
                try:
                    listener.prepare()
                except Exception:
                    # A partially prepared Cheroot instance may already own
                    # a socket; close it before rolling back earlier binds.
                    try:
                        listener.stop()
                    except Exception as exc:
                        logger.warning('Could not close failed website listener (%s).', type(exc).__name__)
                    raise
                self._prepared.append(listener)
        except BaseException:
            self.stop()
            raise

    def serve(self):
        def run_one(listener):
            try:
                listener.serve()
            except Exception as exc:
                if not self._stop.is_set():
                    self._fault = exc
            finally:
                if not self._stop.is_set():
                    self._fault = self._fault or RuntimeError('A website listener stopped unexpectedly.')
                    self._stop.set()

        self._threads = [threading.Thread(target=run_one, args=(listener,), name='RedInkWebListener', daemon=False)
                         for listener in self._listeners]
        for thread in self._threads:
            thread.start()
        self._stop.wait()
        self.stop()
        for thread in self._threads:
            if thread.is_alive():
                thread.join(timeout=15)
        if any(thread.is_alive() for thread in self._threads):
            raise RuntimeError('Website listener did not stop cleanly.')
        if self._fault is not None:
            raise RuntimeError('Website listener failed.') from self._fault

    def stop(self):
        with self._lock:
            if self._stopped:
                return
            self._stopped = True
            self._stop.set()
        for listener in reversed(self._prepared):
            try:
                listener.stop()
            except Exception as exc:
                # Do not silently hide a transport shutdown failure. Cheroot
                # may have already closed a socket; report the type without
                # leaking a document-root or private key path.
                logger.warning('Website listener shutdown raised %s.', type(exc).__name__)


def create_shared_application(admin_app, sites: tuple[StaticWebsite, ...]):
    """Dispatch only explicitly approved prefixes to static sites on the admin port.

    Wrap the WSGI callable for one listener generation instead of mutating the
    Flask admin app. This keeps failed live updates and rollback isolated.
    """
    if not sites:
        return admin_app
    protected = {rule.rule.split('/')[1].casefold() for rule in admin_app.url_map.iter_rules()
                 if rule.rule != '/' and len(rule.rule.split('/')) > 1}
    # Never depend only on a manual route list: newly registered API routes
    # also automatically reserve their top-level namespace.
    protected |= RESERVED_SERVER_URL_ROOTS
    for site in sites:
        if site.path == '/' or site.path.split('/', 2)[1].casefold() in protected:
            raise ConfigError(f"Website {site.name}: URL path overlaps a protected Usage Server route.")
    static_app = create_static_application(sites)
    prefixes = tuple(sorted((site.path for site in sites), key=len, reverse=True))

    def dispatch(environ, start_response):
        path = environ.get('PATH_INFO', '')
        if any(path == prefix or path.startswith(prefix + '/') for prefix in prefixes):
            return static_app(environ, start_response)
        return admin_app(environ, start_response)

    return dispatch


def create_all_listeners(settings: Settings, admin_app):
    """One listener for admin and shared-port sites; others remain isolated."""
    shared = []
    grouped = {}
    for site in settings.websites:
        if not site.enabled:
            continue
        if site.port == settings.port:
            if site.host.casefold() != settings.host.casefold() or site.https != settings.tls_enabled:
                raise ConfigError(f"Website {site.name}: shared port host and protocol must match the Usage Server.")
            shared.append(site)
        else:
            grouped.setdefault((site.host, site.port, site.https), []).append(site)
    main_app = create_shared_application(admin_app, tuple(shared))
    listeners = [create_https_server(settings, main_app) if settings.tls_enabled
                 else create_http_server(settings, main_app)]
    for (host, port, https), sites in sorted(grouped.items()):
        website_app = create_static_application(tuple(sites))
        site_settings = replace(settings, host=host, port=port, tls_enabled=https)
        listeners.append(create_https_server(site_settings, website_app) if https
                         else create_http_server(site_settings, website_app))
    return WebsiteListenerGroup(listeners)
