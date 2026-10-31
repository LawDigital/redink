"""Normalize the actual Red Ink license-counter wire formats.

The collector accepts the v1 marker, report, query, path, JSON POST and form POST
contracts. Optional loadedDates records the sparse client-local dates on which
an add-in was loaded. Legacy timestamps never imply unobserved days.
Inputs are deliberately bounded; incomplete usage
never becomes a billable row. This module does not
validate a commercial licence or infer a person from a machine name.
"""
from __future__ import annotations

import base64
import re
import unicodedata
from collections.abc import Mapping
from datetime import date, datetime, timezone
from urllib.parse import parse_qs, unquote


class ProtocolError(ValueError):
    """A bounded, user-safe explanation of an invalid event (no raw payload)."""


_HOSTS = {"WD": "WD", "WORD": "WD", "EX": "EX", "XL": "EX", "EXCEL": "EX",
          "OL": "OL", "OUTLOOK": "OL", "UK": "UNKNOWN", "UNKNOWN": "UNKNOWN"}
_HASH_PATTERNS = {"network_key": r"n[0-9a-f]{16}", "user_key": r"u[0-9a-f]{32}",
                  "billing_event_id": r"b[0-9a-f]{32}", "host_event_id": r"e[0-9a-f]{32}"}
_CONTROLS = re.compile(r"[\x00-\x1f\x7f]")


def _get(data: Mapping, *names: str, default=""):
    """Ignore field-name case and accept one-element form/query sequences."""
    found = []
    for key, value in data.items():
        if not isinstance(key, str):
            raise ProtocolError("Event field names must be strings.")
        if key.casefold() in {n.casefold() for n in names}:
            if isinstance(value, (list, tuple)) and names[0] not in {"allowedNetworkIds", "loadedDates", "loadedHosts"}:
                if len(value) != 1:
                    raise ProtocolError(f"Field {names[0]} must occur exactly once.")
                value = value[0]
            found.append(value)
    if not found:
        return default
    # Comparing text permits form '1' beside JSON 1, but not two different IDs.
    def comparable(value):
        if isinstance(value, (str, int, bool)) or value is None:
            return str(value)
        if isinstance(value, (list, tuple)) and len(value) <= 100 and all(isinstance(v, str) for v in value):
            return tuple(value)
        raise ProtocolError(f"Field {names[0]} has an invalid value type.")
    if len(found) > 1 and any(comparable(v) != comparable(found[0]) for v in found[1:]):
        raise ProtocolError(f"Conflicting aliases for {names[0]}.")
    return found[0]


def _text(value, field: str, limit: int = 512, *, required: bool = False) -> str:
    if value is None:
        value = ""
    if not isinstance(value, (str, int)) or isinstance(value, bool):
        raise ProtocolError(f"Field {field} must be text.")
    text = unicodedata.normalize("NFC", str(value)).strip()
    if len(text) > limit or _CONTROLS.search(text) or any(unicodedata.category(c) == "Cs" for c in text):
        raise ProtocolError(f"Field {field} exceeds its limit or contains control characters.")
    if required and not text:
        raise ProtocolError(f"Missing required field: {field}.")
    return text


def _integer(value, field: str, low: int, high: int) -> int:
    if not isinstance(value, (str, int)) or isinstance(value, bool) or not re.fullmatch(r"[0-9]{1,9}", str(value)):
        raise ProtocolError(f"Field {field} must be an integer.")
    result = int(value)
    if not low <= result <= high:
        raise ProtocolError(f"Field {field} must be between {low} and {high}.")
    return result


def _boolean(value, field: str) -> bool:
    if isinstance(value, bool):
        return value
    if not isinstance(value, (str, int)):
        raise ProtocolError(f"Field {field} must be true, false, 1 or 0.")
    if str(value).casefold() in {"1", "true"}:
        return True
    if str(value).casefold() in {"0", "false"}:
        return False
    raise ProtocolError(f"Field {field} must be true, false, 1 or 0.")


def _month(value, *, required: bool = True) -> str:
    result = _text(value, "month", 7, required=required)
    if not result and not required:
        return ""
    if not re.fullmatch(r"[0-9]{4}-(?:0[1-9]|1[0-2])", result) or result.startswith("0000"):
        raise ProtocolError("Month must be yyyy-MM.")
    return result


def _hash(value, field: str) -> str:
    result = _text(value, field, 65, required=True).lower()
    if not re.fullmatch(_HASH_PATTERNS[field], result):
        raise ProtocolError(f"Invalid {field} format.")
    return result


def normalize_utc(value, field: str, *, required: bool = False) -> str:
    text = _text(value, field, 64, required=required)
    if not text:
        return ""
    try:
        stamp = datetime.fromisoformat(text.replace("Z", "+00:00").replace("z", "+00:00"))
        if stamp.tzinfo is None:
            raise ValueError("timezone required")
        return stamp.astimezone(timezone.utc).isoformat(timespec="microseconds").replace("+00:00", "Z")
    except (ValueError, OverflowError):
        raise ProtocolError(f"Field {field} must be an ISO timestamp with timezone.") from None


def _list_values(value, field: str, limit: int) -> list[str]:
    """JSON arrays and comma-separated query/form values share one contract."""
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, (list, tuple)) or len(value) > limit or not all(isinstance(item, str) for item in value):
        raise ProtocolError(f"Field {field} must be a list of at most {limit} values.")
    parts = [part.strip() for item in value for part in re.split(r"[,|]", _text(item, field, 512)) if part.strip()]
    if len(parts) > limit:
        raise ProtocolError(f"Field {field} exceeds {limit} values.")
    return sorted(set(parts))


def _loaded_dates(value, month: str) -> list[str]:
    days = _list_values(value, "loadedDates", 31)
    for day in days:
        try:
            if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", day) or date.fromisoformat(day).isoformat() != day or day[:7] != month:
                raise ValueError()
        except ValueError:
            raise ProtocolError("loadedDates must contain real yyyy-MM-dd dates within the event month.") from None
    return days


def _loaded_hosts(value, representative: str) -> list[str]:
    hosts = _list_values(value, "loadedHosts", 4)
    if any(h.upper() not in _HOSTS for h in hosts):
        raise ProtocolError("loadedHosts contains an unrecognized Office host ID.")
    return sorted({representative} | {_HOSTS[h.upper()] for h in hosts})


def safe_product_id(product_id: str) -> str:
    """Match .NET Char iteration, including two UTF-16 units for astral text."""
    normalized = unicodedata.normalize("NFC", product_id).strip().replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()
    units = normalized.encode("utf-16-le", errors="surrogatepass")
    characters = (chr(int.from_bytes(units[index:index + 2], "little")) for index in range(0, len(units), 2))
    value = "".join(c if unicodedata.category(c).startswith("L") or
                    unicodedata.category(c) == "Nd" or c in "-_." else "_"
                    for c in characters)[:80]
    return value or "UNKNOWN_PRODUCT"


def normalize_event(data: Mapping, now: datetime | None = None) -> dict:
    """Return a canonical snake_case record, or raise ProtocolError.

    ``now`` is accepted for callers' test clocks. Freshness and monthly
    acceptance policy belong to the store, not the wire parser.
    Usage IDs must be present because URL-safe product identifiers cannot
    reconstruct the original product ID used by the add-in hash algorithm.
    """
    if not isinstance(data, Mapping) or len(data) > 80:
        raise ProtocolError("Event must be an object with at most 80 fields.")
    version = _integer(_get(data, "v", "ri_lc_v", "version", default=1), "version", 1, 1)
    kind = _text(_get(data, "t", "event_type", default="U"), "t", 1, required=True).upper()
    if kind not in {"U", "E"}:
        raise ProtocolError("Event type must be U or E.")
    result = {"version": version, "event_type": kind}
    if kind == "E":
        result.update(
            month=_month(_get(data, "relatedMonth", "m", "month"), required=False),
            error_event_id=_text(_get(data, "errorEventId", "id", "error_event_id"), "errorEventId", 100, required=True),
            error_code=_text(_get(data, "errorCode", "ec", "error_code"), "errorCode", 100, required=True),
            related_file=_text(_get(data, "relatedFile", "f", "related_file"), "relatedFile", 512),
            detail=_text(_get(data, "detail"), "detail", 2048),
            created_utc=normalize_utc(_get(data, "createdUtc", "created_utc"), "createdUtc"))
        return result

    month = _month(_get(data, "month", "m"))
    network = _hash(_get(data, "networkKey", "n", "network_key"), "network_key")
    user = _hash(_get(data, "userKey", "u", "user_key"), "user_key")
    product = _text(_get(data, "productId", "p", "product_id", default=_get(data, "productIdSafe", "product_id_safe")),
                    "productId", 256, required=True)
    product_safe = _text(_get(data, "productIdSafe", "product_id_safe", default=safe_product_id(product)), "productIdSafe", 80, required=True)
    host_raw = _text(_get(data, "hostId", "h", "host_id"), "hostId", 16, required=True).upper()
    if host_raw not in _HOSTS:
        raise ProtocolError("Unrecognized Office host ID.")
    host = _HOSTS[host_raw]
    scope = _text(_get(data, "networkScope", "ns", "network_scope"), "networkScope", 4096)
    networks = _get(data, "allowedNetworkIds", "ani", "allowed_network_ids", default=[])
    if isinstance(networks, str):
        networks = networks.split("|")
    if not isinstance(networks, (list, tuple)) or len(networks) > 100:
        raise ProtocolError("allowedNetworkIds must be a list of at most 100 network names.")
    if not networks and scope:
        networks = scope.split("|")
    # parse_qs represents ani=one%7Ctwo as a one-element sequence; the
    # network separator has the same meaning in all supported transports.
    network_parts = [part for n in networks for part in _text(n, "allowedNetworkIds", 4096).split("|")]
    if len(network_parts) > 100:
        raise ProtocolError("allowedNetworkIds exceeds 100 network names.")
    normalized_networks = sorted({_text(n, "allowedNetworkIds", 256) for n in network_parts} - {""}, key=str.casefold)
    anonymous = _boolean(_get(data, "anon", "a", "anonymous", default=True), "anonymous")
    user_obf = "" if anonymous else _text(_get(data, "userObf", "du", "user_obf"), "userObf", 2048)
    # This is reversible obfuscation in v1, never an authentication credential.
    if user_obf:
        if not user_obf.lower().startswith("o1_") or not re.fullmatch(r"[A-Za-z0-9_-]+", user_obf[3:]):
            raise ProtocolError("Unsupported user obfuscation format.")
        try:
            decoded = base64.b64decode(user_obf[3:] + "=" * (-len(user_obf[3:]) % 4), altchars=b"-_", validate=True).decode("utf-8")
            _text(decoded, "decoded user identity", 512, required=True)
        except (ValueError, UnicodeError):
            raise ProtocolError("Invalid obfuscated user identity.") from None
    billing = _get(data, "billingEventId", "b", "billing_event_id")
    event_id = _get(data, "hostEventId", "e", "host_event_id")
    if not billing or not event_id:
        raise ProtocolError("Usage events require billingEventId and hostEventId; safe product names cannot reconstruct them.")
    result.update(month=month, network_key=network, network_scope=scope,
                  allowed_network_ids=normalized_networks, product_id=product,
                  product_id_safe=product_safe, host_id=host, user_key=user,
                  billing_event_id=_hash(billing, "billing_event_id"),
                  host_event_id=_hash(event_id, "host_event_id"),
                  first_use_utc=normalize_utc(_get(data, "firstUseUtc", "first_use_utc"), "firstUseUtc"),
                  last_use_utc=normalize_utc(_get(data, "lastUseUtc", "last_use_utc"), "lastUseUtc"),
                  reported_utc=normalize_utc(_get(data, "reportedUtc", "reported_utc"), "reportedUtc"),
                  report_interval_minutes=_integer(_get(data, "reportIntervalMinutes", "report_interval_minutes", default=0), "reportIntervalMinutes", 0, 10080),
                  loaded_dates=_loaded_dates(_get(data, "loadedDates", "loaded_dates", default=[]), month),
                  loaded_hosts=_loaded_hosts(_get(data, "loadedHosts", "loaded_hosts", default=[]), host),
                  anonymous=anonymous, user_obf=user_obf)
    return result


_PATH_RE = re.compile(
    r"(?:^|/)RI_LC_v1/m/(?P<m>[^/\s?#\"]+)/n/(?P<n>[^/\s?#\"]+)"
    r"(?:/ns/(?P<ns>[^/\s?#\"]*))?/p/(?P<p>[^/\s?#\"]+)/h/(?P<h>[^/\s?#\"]+)"
    r"/u/(?P<u>[^/\s?#\"]+)/b/(?P<b>[^/\s?#\"]+)/e/(?P<e>[^/\s?#\"]+)"
    r"/a/(?P<a>[01])(?:/r/(?P<r>[^/\s?#\"]+))?(?:/du/(?P<du>[^/\s?#\"]+))?", re.IGNORECASE)


def parse_legacy_path(path: str) -> dict | None:
    """Parse a v1 path from a request URL/path or an access-log line."""
    if len(path) > 65536:
        raise ProtocolError("Legacy path exceeds 64 KiB.")
    match = _PATH_RE.search(path)
    if not match:
        if "ri_lc_v1/" in path.casefold():
            raise ProtocolError("Malformed v1 license-counter path.")
        return None
    # A match must terminate at the request boundary, not a valid-looking prefix.
    tail = path[match.end():]
    if tail and tail[0] not in " ?#\t\r\n\"":
        raise ProtocolError("Unexpected trailing path segments.")
    try:
        values = {k: unquote(v, encoding="utf-8", errors="strict") for k, v in match.groupdict().items() if v is not None}
    except UnicodeError:
        raise ProtocolError("Legacy path contains invalid UTF-8 URL encoding.") from None
    values.update(ri_lc_v=1, t="U")
    values["productIdSafe"] = values["p"]
    return normalize_event(values)


def parse_legacy_query(text: str) -> dict | None:
    """Parse a v1 query or URL from common/IIS access-log text.

    IIS uses '-' for the empty URI query, which is deliberately ignored. The
    entire query is decoded with form semantics, including '+' spaces.
    """
    if len(text) > 65536:
        raise ProtocolError("Legacy query exceeds 64 KiB.")
    match = re.search(r"ri_lc_v=1(?:&[^\s\"#]*)?", text, re.IGNORECASE)
    if not match:
        return None
    try:
        values = parse_qs(match.group(), keep_blank_values=True, strict_parsing=False,
                          encoding="utf-8", errors="strict", max_num_fields=80)
    except (ValueError, UnicodeError):
        raise ProtocolError("Malformed v1 query.") from None
    return normalize_event(values)
