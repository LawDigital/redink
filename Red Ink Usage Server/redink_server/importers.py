"""Bounded, read-only imports of existing counter files and web access logs.

Invalid records are yielded explicitly. Corrupt marker JSON is never silently
reconstructed from a filename. Callers decide persistence/cursor policy; the
fingerprint contains no original text and is useful for diagnostics/deduplication.
"""
from __future__ import annotations

import hashlib
import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

from .protocol import ProtocolError, normalize_event, parse_legacy_path, parse_legacy_query

DEFAULT_MAX_FILE_BYTES = 16 * 1024 * 1024
DEFAULT_MAX_LINE_BYTES = 65536


@dataclass(frozen=True)
class ImportRecord:
    event: dict | None
    line: int
    error: str | None = None
    skipped: bool = False
    fingerprint: str = ""


def _is_link(path: Path) -> bool:
    """Also refuse Windows junctions/reparse points, not only POSIX symlinks."""
    try:
        info = path.lstat()
        return path.is_symlink() or bool(getattr(info, "st_file_attributes", 0) & 0x400)
    except OSError:
        return False


def safe_iter_files(root: str | Path, *, recursive: bool = True,
                    suffixes: tuple[str, ...] = (".json", ".jsonl", ".log", ".txt", ".csv")) -> Iterator[Path]:
    """Enumerate only ordinary files below a configured root, never symlinks.

    Explicit files are supported. An inaccessible/missing root is an error, so
    the UI can distinguish an outage from a successfully scanned empty folder.
    Symlinks/reparse points are not followed during directory scans.
    """
    root = Path(root)
    if _is_link(root):
        raise OSError("Import roots must not be symbolic links.")
    root = root.resolve(strict=True)
    if root.is_file():
        yield root
        return
    if not root.is_dir():
        raise OSError("Import root is not a regular file or directory.")
    suffix_set = {s.casefold() for s in suffixes}
    def fail(error):
        raise error
    for directory, dirs, files in os.walk(root, followlinks=False, onerror=fail):
        dirs[:] = sorted(d for d in dirs if not _is_link(Path(directory) / d)) if recursive else []
        for name in sorted(files):
            path = Path(directory) / name
            if _is_link(path) or path.suffix.casefold() not in suffix_set:
                continue
            try:
                resolved = path.resolve(strict=True)
                if not resolved.is_relative_to(root) or not resolved.is_file():
                    continue
            except OSError:
                continue
            yield resolved


def _object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ProtocolError("Duplicate JSON field.")
        result[key] = value
    return result


def _fingerprint(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def _normalize(value, line: int, fingerprint: str) -> ImportRecord:
    try:
        event = normalize_event(value)
        return ImportRecord(event, line, fingerprint=fingerprint)
    except ProtocolError as exc:
        return ImportRecord(None, line, str(exc), fingerprint=fingerprint)


def iter_counter_records(path: str | Path, *, max_file_bytes: int = DEFAULT_MAX_FILE_BYTES,
                         max_line_bytes: int = DEFAULT_MAX_LINE_BYTES) -> Iterator[ImportRecord]:
    """Read a marker/report JSON, JSON array, append JSONL or web log.

    UTF-8 with or without BOM matches the add-in writer. A trailing partial
    append line is reported as invalid and will be retried on the next scan.
    Files larger than the configured cap are explicitly rejected, never
    partially treated as complete. Increase the cap or rotate source logs.
    """
    path = Path(path)
    if _is_link(path):
        yield ImportRecord(None, 0, "Symbolic-link input refused.")
        return
    try:
        stat = path.stat()
        if stat.st_size > max_file_bytes:
            yield ImportRecord(None, 0, f"File exceeds the configured {max_file_bytes}-byte import limit.")
            return
        if path.name.casefold() == "ri_lc_state.json":
            yield ImportRecord(None, 0, skipped=True)
            return
        if path.suffix.casefold() == ".json":
            with path.open("rb") as stream:
                raw = stream.read(max_file_bytes + 1)
            if len(raw) > max_file_bytes:
                yield ImportRecord(None, 0, "File grew beyond the import limit while being read.")
                return
            digest = _fingerprint(raw)
            try:
                value = json.loads(raw.decode("utf-8-sig"), object_pairs_hook=_object)
            except (UnicodeError, ValueError, RecursionError) as exc:
                message = str(exc) if isinstance(exc, ProtocolError) else "Invalid UTF-8 JSON; no filename fallback was used."
                yield ImportRecord(None, 1, message, fingerprint=digest)
                return
            if isinstance(value, list):
                if len(value) > 100000:
                    yield ImportRecord(None, 0, "JSON event array exceeds 100000 records.")
                    return
                for number, item in enumerate(value, 1):
                    yield _normalize(item, number, _fingerprint(json.dumps(item, sort_keys=True, ensure_ascii=True).encode("ascii")))
            else:
                yield _normalize(value, 1, digest)
            return
        total = 0
        with path.open("rb") as stream:
            line_number = 0
            while True:
                raw = stream.readline(max_line_bytes + 1)
                if not raw:
                    break
                total += len(raw)
                line_number += 1
                if total > max_file_bytes:
                    yield ImportRecord(None, line_number, "File grew beyond the import limit while being read.")
                    return
                digest = _fingerprint(raw)
                if len(raw) > max_line_bytes:
                    yield ImportRecord(None, line_number, "Line exceeds the configured import limit.", fingerprint=digest)
                    # Do not interpret the remainder as separate records.
                    while raw and not raw.endswith(b"\n"):
                        raw = stream.readline(max_line_bytes + 1)
                        total += len(raw)
                        if total > max_file_bytes:
                            return
                    continue
                try:
                    line = raw.decode("utf-8-sig" if line_number == 1 else "utf-8").strip()
                except UnicodeError:
                    yield ImportRecord(None, line_number, "Invalid UTF-8 line.", fingerprint=digest)
                    continue
                if not line or line.startswith("#"):
                    yield ImportRecord(None, line_number, skipped=True, fingerprint=digest)
                    continue
                if line.startswith("{"):
                    try:
                        value = json.loads(line, object_pairs_hook=_object)
                    except (ValueError, UnicodeError, RecursionError):
                        yield ImportRecord(None, line_number, "Invalid JSON line; it will be retried when the file changes.", fingerprint=digest)
                        continue
                    yield _normalize(value, line_number, digest)
                    continue
                try:
                    event = parse_legacy_path(line) or parse_legacy_query(line)
                    if event:
                        yield ImportRecord(event, line_number, fingerprint=digest)
                    else:
                        yield ImportRecord(None, line_number, skipped=True, fingerprint=digest)
                except (ProtocolError, UnicodeError) as exc:
                    yield ImportRecord(None, line_number, str(exc) if isinstance(exc, ProtocolError) else "Invalid URL encoding.", fingerprint=digest)
    except OSError:
        yield ImportRecord(None, 0, "File could not be opened or read; check path and service-account permissions.")
