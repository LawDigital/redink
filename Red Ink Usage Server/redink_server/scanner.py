"""Background imports retain committed history when sources are unavailable."""
from __future__ import annotations

import hashlib
import logging
import threading
from collections import Counter
from pathlib import Path

from .config import Settings
from .importers import iter_counter_records, safe_iter_files
from .log_parser import iter_log_records, iter_log_delta, log_host_from_name, logical_log_name, LogFraming
from .protocol import ProtocolError
from .store import Store

logger = logging.getLogger(__name__)


class Scanner:
    def __init__(self, settings: Settings, store: Store):
        self.settings = settings
        self.store = store
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._wake = threading.Event()
        self._thread = None

    @staticmethod
    def _log_signatures(path: Path, offset: int = 0) -> tuple[str, str]:
        """Compare first line and committed tail to detect log replacement."""
        with path.open("rb") as stream:
            head = hashlib.sha256(stream.readline(65537)).hexdigest()
            tail = ""
            if offset:
                stream.seek(max(0, offset - 128))
                tail = hashlib.sha256(stream.read(min(offset,128))).hexdigest()
        return head, tail

    def _scan_log_file(self, path: Path, stat, cached, force: bool):
        cursor = self.store.log_scan_cursor(path)
        head, current_tail = self._log_signatures(path, cursor["offset_bytes"] if cursor else 0)
        resumed = bool(cursor and not force and stat.st_size >= cursor["offset_bytes"]
                       and cursor["head_signature"] in ("fr2:" + head + ":0", "fr2:" + head + ":1") and cursor["tail_signature"] == current_tail)
        framing = LogFraming(resumed and cursor['head_signature'].endswith(':1'))
        offset = cursor["offset_bytes"] if resumed else 0
        line_no = cursor["line_number"] if resumed else 0
        if (resumed and cached and offset >= stat.st_size and
                (cached["size"],cached["mtime_ns"]) == (stat.st_size,stat.st_mtime_ns)
                and "Incomplete final" not in cached["detail"]):
            return cached["records"],cached["errors"],cached["skipped"],cached["detail"]
        records = (cached["records"] if resumed and cached else 0)
        errors = (cached["errors"] if resumed and cached else 0)
        skipped = (cached["skipped"] if resumed and cached else 0)
        if resumed and cached and "Incomplete final" in cached["detail"]:
            errors = max(0, errors - 1)
        details = []
        new_rows = []
        legacy_keys = []
        legacy_occurrences = Counter()
        legacy_upgrade = cursor is None and cached is not None and cached["records"] > 0
        committed_diagnostic = framing.diagnostic
        new_offset = offset
        new_line = line_no
        for item,end_offset in iter_log_delta(path, start_offset=offset, framing=framing,
                        max_scan_bytes=max(65536,self.settings.maximum_file_bytes)):
            if self._stop.is_set():
                details.append("Scan interrupted; next scan resumes at the last committed position.")
                break
            if item.error:
                errors += 1
                if len(details) < 5: details.append(f"Line {line_no+item.line}: {item.error}")
            elif item.skipped:
                skipped += 1
            elif item.event:
                # A stable line index survives BOM/CRLF conversion and copying
                # the same log to another root. Identical repeated lines at
                # different indices are distinct invocations.
                key = hashlib.sha256((logical_log_name(path) + "|" + item.event["host_id"] + "|"
                        + item.fingerprint + "|" + str(line_no+item.line)).encode()).hexdigest()
                new_rows.append((key,item.event))
                if legacy_upgrade:
                    legacy_occurrences[item.fingerprint] += 1
                    old = hashlib.sha256((logical_log_name(path)+"|"+item.event["host_id"]+"|"+item.fingerprint+"|"+str(legacy_occurrences[item.fingerprint])).encode()).hexdigest()
                    legacy_keys.append((old,key,item.event))
                records += 1
            if end_offset > new_offset:
                committed_diagnostic = framing.diagnostic
                new_offset = end_offset
                new_line = line_no + item.line
            if end_offset == offset and item.error and "Incomplete" in item.error:
                break
        try:
            after = path.stat()
            if (after.st_size,after.st_mtime_ns) != (stat.st_size,stat.st_mtime_ns):
                details.append("Source changed during scan; new content will be checked on the next scan.")
            if new_offset < after.st_size and not any("Incomplete" in x for x in details):
                details.append("Additional log data is pending; the next scheduled scan will continue.")
            _, tail = self._log_signatures(path,new_offset)
        except OSError:
            errors += 1
            details.append("Source disappeared during scan; stored history is retained.")
            tail = current_tail
        # A completed newline-terminated chunk and its cursor are committed
        # together, so a crash never marks uncommitted rows as consumed.
        detail = " ".join(details)
        self.store.save_log_progress(path,stat,rows=new_rows,legacy_keys=legacy_keys,offset=new_offset,line_number=new_line,
            head="fr2:" + head + (":1" if committed_diagnostic else ":0"),tail=tail,records=records,errors=errors,skipped=skipped,detail=detail)
        return records, errors, skipped, detail

    def scan_path(self, root: Path, kind: str, *, recursive: bool | None = None, force: bool = False, allow_recent: bool = False) -> dict:
        if kind not in {"counter", "log"}:
            raise ValueError("Import kind must be counter or log.")
        with self._lock:
            total = {"kind": kind, "path": str(root), "files": 0, "records": 0, "errors": 0, "skipped": 0, "detail": ""}
            messages = []
            try:
                for path in safe_iter_files(root, recursive=self.settings.recursive if recursive is None else recursive):
                    if kind == "log" and log_host_from_name(path) is None:
                        continue  # Not a usage source: no cursor, file budget, parsing or pending warning.
                    total["files"] += 1
                    if total["files"] > self.settings.maximum_files_per_scan:
                        total["errors"] += 1
                        messages.append("File-count limit reached; split the source or increase maximum_files_per_scan.")
                        break
                    try:
                        stat = path.stat()
                    except OSError:
                        total["errors"] += 1
                        messages.append("A source file became unavailable during the scan.")
                        continue
                    cached = self.store.file_status(kind, path)
                    if kind == "log":
                        records, errors, skipped, detail = self._scan_log_file(path,stat,cached,force)
                        total["records"] += records
                        total["errors"] += errors
                        total["skipped"] += skipped
                        if detail and len(messages) < 5:
                            messages.append(path.name + ": " + detail)
                        continue
                    if not force and cached and cached["size"] == stat.st_size and cached["mtime_ns"] == stat.st_mtime_ns and not cached["errors"]:
                        for key in ("records", "errors", "skipped"):
                            total[key] += cached[key]
                        continue
                    records = errors = skipped = 0
                    details = []
                    occurrences = Counter()
                    batch = []
                    iterator = iter_counter_records if kind == "counter" else iter_log_records
                    for item in iterator(path, max_file_bytes=self.settings.maximum_file_bytes):
                        if self._stop.is_set():
                            errors += 1
                            details.append("Scan interrupted during server shutdown; next scan will retry.")
                            break
                        if item.error:
                            errors += 1
                            if len(details) < 5:
                                details.append(f"Line {item.line}: {item.error}")
                        elif item.skipped:
                            skipped += 1
                        elif item.event:
                            if kind == "counter":
                                try:
                                    self.store.ingest(item.event, source="counter file: " + str(path), live_request=allow_recent)
                                    records += 1
                                except ProtocolError as exc:
                                    errors += 1
                                    if len(details) < 5:
                                        details.append(f"Line {item.line}: {exc}")
                            else:
                                # Same copied or rotated log has the same IDs.
                                # Repeated identical lines within a file keep their
                                # occurrence number (two real calls may share a ms).
                                occurrences[item.fingerprint] += 1
                                key = hashlib.sha256((logical_log_name(path) + "|" + item.event["host_id"] + "|" + item.fingerprint + "|" + str(occurrences[item.fingerprint])).encode()).hexdigest()
                                batch.append((key, item.event))
                                records += 1
                                if len(batch) >= 500:
                                    self.store.add_log_records(batch, str(path))
                                    batch.clear()
                    if batch:
                        self.store.add_log_records(batch, str(path))
                    # If a writer appended while scanning, never cache a partial
                    # observation as the final version. It will be reread next time.
                    try:
                        after = path.stat()
                        if (after.st_size, after.st_mtime_ns) != (stat.st_size, stat.st_mtime_ns):
                            errors += 1
                            details.append("Source changed during the scan; automatically retried next scan.")
                    except OSError:
                        errors += 1
                        details.append("Source disappeared during the scan; committed records were retained.")
                    detail = " ".join(details)
                    self.store.save_file_status(kind, path, stat, records, errors, skipped, detail)
                    total["records"] += records
                    total["errors"] += errors
                    total["skipped"] += skipped
                    if detail and len(messages) < 5:
                        messages.append(path.name + ": " + detail)
            except OSError:
                total["errors"] += 1
                messages.append("Source is unavailable. Check the path and the server account's permissions; stored history was retained.")
            total["detail"] = " ".join(messages)[:2000]
            self.store.source_status(kind, root, total["records"], total["errors"], total["skipped"], total["detail"])
            return total

    def scan_all(self) -> list[dict]:
        results = []
        for kind, roots in (("counter", self.settings.counter_paths), ("log", self.settings.log_paths)):
            for root in roots:
                if self._stop.is_set():
                    return results
                results.append(self.scan_path(root, kind, allow_recent=True))
        self.store.maintenance()
        return results

    def update_settings(self, settings: Settings, *, scan_now: bool = True):
        """Apply validated, reloadable import settings without restarting the process."""
        self.settings = settings
        if scan_now:
            self._wake.set()

    def _run(self):
        while not self._stop.is_set():
            try:
                self.scan_all()
            except Exception:
                # No raw input, request URLs, credentials or prompt text in logs.
                logger.exception("Background scan failed; the next scheduled scan will retry.")
            self._wake.wait(self.settings.scan_seconds)
            self._wake.clear()

    def start(self):
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._run, name="RedInkImports", daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
        self._wake.set()
        if self._thread is not None:
            self._thread.join(timeout=5)
