"""Metadata-only parser for SharedLogger.vb's actual LogPath format.

These records count feature invocations, not licences, model requests, tokens
or cost. Dates are client-local because the source writes no timezone. User
hashes derive only from USERNAME and must not be joined to licence identities.
"""
from __future__ import annotations

import hashlib
import re
from datetime import datetime
from pathlib import Path
from typing import Iterator

from .importers import DEFAULT_MAX_FILE_BYTES, DEFAULT_MAX_LINE_BYTES, ImportRecord, _is_link

_LOG_RE = re.compile(r"^\[(?P<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3})\] (?P<user>[A-Fa-f0-9]{16}) (?P<rest>.*)$")
_FILE_RE = re.compile(r"^redink-[A-Fa-f0-9]{16}(?:-\d{4}(?:0[1-9]|1[0-2]))?-(WD|XL|EX|OL|UK|UNKNOWN)(?:\s*\(\d+\))?\.log$", re.IGNORECASE)
_VERSION_RE = re.compile(r"V(?:[0-9]|\.[0-9])[A-Za-z0-9_.+\-]{0,63}")
_LOG_VERSION_PREFIX = re.compile(r"^(V(?:[0-9]|\.[0-9])[A-Za-z0-9_.+\-]{0,63})\s+(.+)$")
_BUILD_CHANNEL = re.compile(r"^(?:Develop|Preview|Release|Stable|Beta|Debug)\s+", re.IGNORECASE)
_FEATURE_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.+\-]*(?: \([A-Za-z0-9_.+\- ]{1,60}\))?")
_HOSTS = {"WD": "WD", "XL": "EX", "EX": "EX", "OL": "OL", "UK": "UNKNOWN", "UNKNOWN": "UNKNOWN"}


def log_host_from_name(filename: str | Path) -> str | None:
    match = _FILE_RE.fullmatch(Path(filename).name)
    return _HOSTS[match.group(1).upper()] if match else None


def logical_log_name(path: Path) -> str:
    # A copied download with "(1)" is the same stream, not a second source.
    return re.sub(r"\s*\(\d+\)(?=\.log$)", "", path.name.casefold())


class LogFraming:
    """Remember just one bit (previous header was diagnostic), never its text."""
    def __init__(self, diagnostic=False):
        self.diagnostic = diagnostic

    def line(self, text, host, number, digest):
        if not text.strip():
            return ImportRecord(None, number, skipped=True, fingerprint=digest)
        event = parse_log_line(text, host)
        if event is not None:
            self.diagnostic = not event['invoked']
            return ImportRecord(event, number, fingerprint=digest)
        if self.diagnostic and _CONTINUATION.match(text):
            return ImportRecord(None, number, skipped=True, fingerprint=digest)
        self.diagnostic = False
        return ImportRecord(None, number, 'Unrecognized SharedLogger record; raw text not retained.', fingerprint=digest)


_CONTINUATION = re.compile(r"^(?:[ \t]+(?:at|bei)\s+|[ \t]*--- (?:End of|Ende )|Validated structure window:|PREPARATION SKIPPED:)")
_TOOL_KEY = re.compile(r"AgentToolCall_(?:Outlook_(?:AutoPilot|LocalChatAgent|Tooling)|Word_Agent)_\w{1,150}", re.I)
_STORE_LABEL = re.compile(r"AgentToolCall_(?:Outlook_(?:AutoPilot|LocalChatAgent|Tooling)|Word_Agent)_knowledge_search_store \(([^\r\n]{1,96})\)", re.I)


def parse_log_line(line: str, host: str = "UNKNOWN") -> dict | None:
    """Return only timestamp/date, hashed user, host, version and safe event key.

    Version is a minimum-20-character field. The exact writer is supported for
    empty and longer versions too. General diagnostic text is classified as
    'Other activity' and never retained or returned.
    """
    line = line.lstrip("\ufeff").rstrip("\r\n")
    match = _LOG_RE.fullmatch(line)
    if not match or len(line) > DEFAULT_MAX_LINE_BYTES:
        return None
    try:
        stamp = datetime.strptime(match.group("timestamp"), "%Y-%m-%d %H:%M:%S.%f")
    except ValueError:
        return None
    rest = match.group("rest")
    if len(rest) < 21:
        return None
    # SharedLogger versions may be "V.030126 Develop" or "V.120126 Preview"
    # with padding, or the newer "V2.3.1" token. The build channel is not
    # the feature label. Long version strings are deliberately supported.
    match_version = _LOG_VERSION_PREFIX.match(rest)
    if match_version:
        version = match_version.group(1)
        remainder = _BUILD_CHANNEL.sub("", match_version.group(2).strip(), count=1).strip()
    elif rest[:20].strip() == "":
        version = ""
        remainder = rest[20:].strip()
    else:
        return None
    if not remainder:
        return None
    invoked = bool(re.search(r" invoked\.?$", remainder, re.IGNORECASE))
    label = re.sub(r" invoked\.?$", "", remainder, flags=re.IGNORECASE) if invoked else ""
    feature = label if invoked and (_FEATURE_RE.fullmatch(label) or _TOOL_KEY.fullmatch(label) or (_STORE_LABEL.fullmatch(label) and label.isprintable())) and len(label) <= 200 else ("Other invocation" if invoked else "Other activity")
    return {
        "timestamp_local": stamp.isoformat(timespec="milliseconds"),
        "date": stamp.date().isoformat(),
        "host": _HOSTS.get(host.upper(), "UNKNOWN"),
        "host_id": _HOSTS.get(host.upper(), "UNKNOWN"),
        "user_hash": match.group("user").upper(),
        "version": version,
        "feature": feature,
        "invoked": invoked,
        "is_invoked": invoked,
        "category": "invocation" if invoked else "other",
    }


def iter_log_records(path: str | Path, *, max_file_bytes: int = DEFAULT_MAX_FILE_BYTES,
                     max_line_bytes: int = DEFAULT_MAX_LINE_BYTES) -> Iterator[ImportRecord]:
    path = Path(path)
    host = log_host_from_name(path)
    if host is None:
        yield ImportRecord(None, 0, skipped=True)
        return
    if _is_link(path):
        yield ImportRecord(None, 0, "Symbolic-link input refused.")
        return
    try:
        if path.stat().st_size > max_file_bytes:
            yield ImportRecord(None, 0, f"File exceeds the configured {max_file_bytes}-byte import limit.")
            return
        total = 0
        number = 0
        framing = LogFraming()
        with path.open("rb") as stream:
            while True:
                raw = stream.readline(max_line_bytes + 1)
                if not raw:
                    break
                number += 1
                total += len(raw)
                if total > max_file_bytes:
                    yield ImportRecord(None, number, "File grew beyond the import limit while being read.")
                    return
                digest = hashlib.sha256(raw).hexdigest()
                if len(raw) > max_line_bytes:
                    framing.diagnostic = False
                    yield ImportRecord(None, number, "Line exceeds the configured log import limit.", fingerprint=digest)
                    while raw and not raw.endswith(b"\n"):
                        raw = stream.readline(max_line_bytes + 1)
                        total += len(raw)
                        if total > max_file_bytes:
                            return
                    continue
                try:
                    line = raw.decode("utf-8-sig" if number == 1 else "utf-8")
                except UnicodeError:
                    framing.diagnostic = False
                    yield ImportRecord(None, number, "Invalid UTF-8 log line.", fingerprint=digest)
                    continue
                # SharedLogger.WriteLine always appends a newline. A writer may
                # be observed between writes, so never commit an unfinished tail
                # even if its temporary text already resembles a valid event.
                if not raw.endswith(b"\n"):
                    yield ImportRecord(None, number, "Incomplete final SharedLogger line; it will be retried when the writer finishes.", fingerprint=digest)
                    continue
                # BOM and CRLF are transport details, not invocation identity.
                # This keeps copied/converted logs idempotent without retaining text.
                digest = hashlib.sha256(line.lstrip("\ufeff").rstrip("\r\n").encode("utf-8")).hexdigest()
                yield framing.line(line, host, number, digest)
    except OSError:
        yield ImportRecord(None, 0, "Log file could not be opened or read; check service-account permissions.")


def iter_log_delta(path: str | Path, *, start_offset: int = 0, framing: LogFraming | None = None,
                   max_scan_bytes: int = DEFAULT_MAX_FILE_BYTES,
                   max_line_bytes: int = DEFAULT_MAX_LINE_BYTES):
    """Yield (ImportRecord, next_byte_offset) from a growing SharedLogger file.

    The max_scan_bytes budget is per scan, not a total log-file limit. Only
    newline-terminated records advance the cursor. The caller commits cursor
    and inserted rows together for crash-safe, bounded incremental scans.
    """
    path = Path(path)
    host = log_host_from_name(path)
    if host is None:
        yield ImportRecord(None, 0, skipped=True), start_offset
        return
    if _is_link(path):
        yield ImportRecord(None, 0, "Symbolic-link input refused."), start_offset
        return
    framing = framing or LogFraming()
    if start_offset < 0 or max_scan_bytes < max_line_bytes:
        raise ValueError("Invalid log cursor or scan budget.")
    try:
        with path.open("rb") as stream:
            stream.seek(start_offset)
            total_end = start_offset + max_scan_bytes
            line_number = 0
            while stream.tell() < total_end:
                begin = stream.tell()
                raw = stream.readline(max_line_bytes + 1)
                if not raw:
                    break
                line_number += 1
                if len(raw) > max_line_bytes:
                    framing.diagnostic = False
                    # A large malformed diagnostic line is deliberately skipped;
                    # do not retain or interpret its contents as usage.
                    while raw and not raw.endswith(b"\n"):
                        raw = stream.readline(max_line_bytes + 1)
                    if not raw.endswith(b"\n"):
                        yield ImportRecord(None,line_number,"Incomplete oversized final log line; retry later."),begin
                        break
                    yield ImportRecord(None,line_number,"Line exceeds the configured log import limit."),stream.tell()
                    continue
                if not raw.endswith(b"\n"):
                    yield ImportRecord(None,line_number,"Incomplete final SharedLogger line; retry after the writer finishes."),begin
                    break
                try:
                    line = raw.decode("utf-8-sig" if begin == 0 else "utf-8")
                except UnicodeError:
                    framing.diagnostic = False
                    yield ImportRecord(None,line_number,"Invalid UTF-8 log line."),stream.tell()
                    continue
                digest = hashlib.sha256(line.lstrip("\ufeff").rstrip("\r\n").encode("utf-8")).hexdigest()
                yield framing.line(line, host, line_number, digest), stream.tell()
    except OSError:
        yield ImportRecord(None,0,"Log file could not be opened or read; check permissions."),start_offset
