"""Open only a token-free dashboard URL; never depend on a browser to serve requests."""
import ipaddress
import logging
import re
import sys
import threading
import webbrowser


def dashboard_url(settings):
    host = settings.host.strip('[]')
    if host in ('0.0.0.0', '::', '*'):
        # A concrete trusted DNS name is useful for a TLS certificate SAN.
        host = next((h.strip('[]') for h in settings.trusted_hosts if h and
                     not h.startswith('.') and h.strip('[]') not in ('localhost','127.0.0.1','::1','0.0.0.0','::')),
                    'localhost' if settings.tls_enabled else ('::1' if host == '::' else '127.0.0.1'))
    try:
        address = ipaddress.ip_address(host)
        if '%' in host:
            raise ValueError('Scoped addresses are not supported for automatic browser launch.')
        host = '[' + str(address) + ']' if address.version == 6 else str(address)
    except ValueError:
        if not re.fullmatch(r'[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?',host):
            raise ValueError('No safe browser hostname; open the configured dashboard manually.') from None
    return f"{'https' if settings.tls_enabled else 'http'}://{host}:{settings.port}/"


def browser_requested(choice, stream=None):
    stream = sys.stdin if stream is None else stream
    return bool(stream and stream.isatty()) if choice is None else choice


def open_after_bind(url):
    def launch():
        try:
            if webbrowser.open_new_tab(url):
                return
        except Exception:
            pass
        logging.getLogger(__name__).warning('No browser opened. Open %s manually.', url)
    threading.Thread(target=launch, name='RedInkBrowser', daemon=True).start()
