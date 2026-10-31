"""Synthetic source-faithful protocol fixtures; no customer or production data."""
from __future__ import annotations
import hashlib
import unicodedata
from redink_server.protocol import safe_product_id


def _hex(value: str, length: int) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()[:length]


def usage(month="2026-10", host="WD", user="alice", product="Red Ink Pro", scope="EXAMPLE.ORG", last_use=None, reported=None, interval_minutes=0, loaded_dates=None, loaded_hosts=None):
    """Return a v1 wire dictionary with exact add-in hash namespaces and IDs."""
    network_ids = sorted(set(scope.split("|")), key=str.casefold)
    scope = "|".join(network_ids)
    user = unicodedata.normalize("NFC", user).strip().replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()
    network_key = "n" + _hex("RI-LC-JNET-v1|" + scope, 16)
    # Matches the provided build's empty LicenseCounterSalt, without embedding a secret.
    user_key = "u" + _hex("RI-LC-USER-v1|" + network_key + "||" + user, 32)
    billing_id = "b" + _hex("RI-LC-BILL-v1|" + "|".join((month, network_key, product, user_key)), 32)
    host = {"XL": "EX", "EXCEL": "EX", "WORD": "WD", "OUTLOOK": "OL"}.get(host.upper(), host.upper())
    host_id = "e" + _hex("RI-LC-HOST-v1|" + "|".join((month, network_key, product, host, user_key)), 32)
    result = {
        "v": 1, "t": "U", "month": month,
        "networkKey": network_key, "networkScope": scope, "allowedNetworkIds": network_ids,
        "productId": product, "productIdSafe": safe_product_id(product), "hostId": host,
        "userKey": user_key, "billingEventId": billing_id, "hostEventId": host_id,
        "firstUseUtc": month + "-01T12:00:00.0000000Z", "anon": True,
        "reportIntervalMinutes": interval_minutes,
    }
    if last_use is not None:
        result["lastUseUtc"] = last_use
    if reported is not None:
        result["reportedUtc"] = reported
    if loaded_dates is not None:
        result["loadedDates"] = loaded_dates
    if loaded_hosts is not None:
        result["loadedHosts"] = loaded_hosts
    return result


def log_line(message="Correct_Outlook invoked", version="V2.3.1", user="alice", timestamp="2026-10-30 14:15:16.123"):
    """Exact SharedLogger writer layout; timestamp deliberately has no timezone."""
    user_hash = _hex(user, 16).upper()
    return f"[{timestamp}] {user_hash} {version:<20} {message}\n"


def log_filename(user="alice", machine="SYNTHETIC-PC", host="WD"):
    host = "XL" if host.upper() in {"EX", "XL", "EXCEL"} else host.upper()
    file_hash = _hex("|".join((user, machine, host)), 16).upper()
    return f"redink-{file_hash}-{host}.log"
