"""Verified monthly log bundles and bounded, licence-independent log storage.

Capture closed segments before verification. Commit and verify the complete
replacement archive before unlinking captured sources. Retain sources/hold files
on every failed stage. The normal import scanner never truncates source files.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import date
from hashlib import sha256
import json
import logging
import os
from pathlib import Path
import re
import shutil
import tempfile
from uuid import uuid4
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from .importers import safe_iter_files, DEFAULT_MAX_LINE_BYTES
from .log_parser import LogFraming, log_host_from_name, logical_log_name

_SEGMENT = re.compile(r'redink-[0-9a-f]{16}-(\d{4})(0[1-9]|1[0-2])-(WD|XL|EX|OL|UK)\.log', re.I)
_BUNDLE = re.compile(r'redink-(\d{4}-\d{2})\.zip')
_MANIFEST = '_redink_archive.json'
logger = logging.getLogger(__name__)


def _digest(stream):
    h = sha256()
    for block in iter(lambda: stream.read(1024 * 1024), b''):
        h.update(block)
    return h.hexdigest()


def month_boundary(keep_months, today=None):
    """Keep current month and the previous N-1 complete calendar months."""
    if not 1 <= keep_months <= 1200:
        raise ValueError('keep-months must be between 1 and 1200.')
    now = today or date.today()
    year, month = divmod(now.year * 12 + now.month - keep_months, 12)
    return f'{year:04}-{month+1:02}'


def _destination(store, destination):
    dest = Path(destination).resolve()
    protected = (*store.settings.log_paths, *store.settings.counter_paths, store.settings.config_directory)
    if any(dest.is_relative_to(Path(root).resolve()) for root in protected):
        raise ValueError('Use an archive directory outside publishing and import roots.')
    return dest


@contextmanager
def _archive_lock(dest):
    """Lifetime OS lock, shared by scheduled maintenance and explicit CLI archival."""
    dest.mkdir(mode=0o700, parents=True, exist_ok=True)
    with (dest / '.redink-archive.lock').open('a+b') as stream:
        if stream.tell() == 0:
            stream.write(b'0'); stream.flush()
        stream.seek(0)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == 'nt':
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def _verified_manifest(z, month):
    if len(z.namelist()) != len(set(z.namelist())):
        raise ValueError('Duplicate archive entries; keep archive and sources.')
    info = z.getinfo(_MANIFEST)
    if info.file_size > 16 * 1024 * 1024:
        raise ValueError('Archive manifest exceeds the verification limit.')
    manifest = json.loads(z.read(_MANIFEST))
    if (not isinstance(manifest, dict) or manifest.get('format') != 'redink-log-archive-v1' or manifest.get('month') != month
            or not isinstance(manifest.get('entries'), dict)
            or set(z.namelist()) != set(manifest['entries']) | {_MANIFEST}):
        raise ValueError('Archive manifest does not match its entries.')
    for name, evidence in manifest['entries'].items():
        if (not isinstance(evidence, dict) or not re.fullmatch(r'[0-9a-f]{64}', evidence.get('sha256', ''))
                or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', evidence.get('newest_date', ''))):
            raise ValueError('Invalid archive evidence.')
        with z.open(name) as entry:
            if _digest(entry) != evidence['sha256']:
                raise ValueError('Archive byte verification failed.')
    return manifest


def _source_evidence(store, path, source, stat, status, *, allow_expired):
    framing = LogFraming(); seen = 0; original = sha256(); length = 0; newest = '0001-01-01'
    cutoff = store.log_retention_cutoff() if allow_expired else ''
    with source.open('rb') as stream, store.connect() as con:
        con.execute('BEGIN')
        for number, raw in enumerate(iter(lambda: stream.readline(DEFAULT_MAX_LINE_BYTES + 1), b''), 1):
            if len(raw) > DEFAULT_MAX_LINE_BYTES or not raw.endswith(b'\n'):
                raise ValueError('Incomplete or oversized record.')
            length += len(raw); original.update(raw)
            text = raw.decode('utf-8-sig' if number == 1 else 'utf-8')
            fp = sha256(text.lstrip('\ufeff').rstrip('\r\n').encode()).hexdigest()
            row = framing.line(text, log_host_from_name(path), number, fp)
            if row.error:
                raise ValueError('Record did not pass framing checks.')
            if row.event:
                newest = max(newest, row.event['date'])
                key = sha256(f'{logical_log_name(path)}|{row.event["host_id"]}|{fp}|{number}'.encode()).hexdigest()
                if not (cutoff and row.event['date'] < cutoff):
                    if not con.execute('SELECT 1 FROM log_events WHERE record_id=?', (key,)).fetchone():
                        raise ValueError('Not every retained record is present in SQLite.')
                seen += 1
        if length != stat.st_size or seen != status['records']:
            raise ValueError('Source no longer matches the complete import.')
    return dict(sha256=original.hexdigest(), newest_date=newest)


def _restore(hold, path, result):
    if hold is None or not hold.exists():
        return
    try:
        # Never overwrite a file recreated by a late/offline writer.
        os.link(hold, path); hold.unlink()
    except OSError:
        result['files'].append({'name': hold.name, 'state': 'manual recovery required'})


def _archive_segments(store, before, dest, *, apply, today, allow_expired):
    result = dict(mode='apply' if apply else 'preview', eligible=0, archived=0, held=0, legacy_files=0, files=[])
    groups = {}; visited = set()
    for root in store.settings.log_paths:
        try:
            for p in safe_iter_files(root, recursive=store.settings.recursive, suffixes=None):
                original, separator, token = p.name.partition('.archive-hold-')
                if separator and _SEGMENT.fullmatch(logical_log_name(Path(original))) and re.fullmatch(r'[0-9a-f]{32}', token):
                    if not apply:
                        continue
                    restored = p.with_name(original)
                    try:
                        os.link(p, restored); p.unlink(); p = restored
                    except OSError:
                        result['held'] += 1
                        result['files'].append(dict(name=p.name, state='manual recovery required',
                            reason='Interrupted archive capture conflicts with a writer or cannot be restored.'))
                        continue
                match = _SEGMENT.fullmatch(logical_log_name(p))
                if not match and log_host_from_name(p) is not None:
                    result['legacy_files'] += 1
                if p in visited or not match or f'{match[1]}-{match[2]}' >= before:
                    continue
                visited.add(p)
                groups.setdefault(f'{match[1]}-{match[2]}', []).append(p)
        except OSError:
            result['held'] += 1
            result['files'].append(dict(name=root.name, state='retained', reason='Log source is unavailable.'))
    # One atomic ZIP replacement per month; never one rewrite per source file.
    for month, paths in sorted(groups.items()):
        captures = []; temporary = None
        try:
            for p in paths:
                hold = None
                try:
                    status = store.file_status('log', p); cursor = store.log_scan_cursor(p); st = p.stat()
                    if (not status or not cursor or status['errors'] or status['size'] != st.st_size
                            or status['mtime_ns'] != st.st_mtime_ns or cursor['offset_bytes'] != st.st_size
                            or not cursor['head_signature'].startswith('fr2:') or not status['records']):
                        raise ValueError('No current, complete, error-free import.')
                    if apply:
                        hold = p.with_name(p.name + '.archive-hold-' + uuid4().hex)
                        p.rename(hold)
                    source = hold or p
                    if source.is_symlink():
                        raise ValueError('Captured source must not be a symbolic link.')
                    evidence = _source_evidence(store, p, source, st, status, allow_expired=allow_expired)
                    result['eligible'] += 1
                    if apply:
                        captures.append((p, hold, evidence))
                        hold = None  # The group owns restoration from here.
                    else:
                        result['files'].append(dict(name=p.name, state='eligible'))
                except (ValueError, OSError, UnicodeError, RuntimeError) as exc:
                    result['held'] += 1
                    result['files'].append(dict(name=p.name, state='retained', reason=str(exc)))
                finally:
                    _restore(hold, p, result)
            if not captures:
                continue
            archive = dest / f'redink-{month}.zip'
            if archive.is_symlink():
                raise ValueError('Archive must not be a symbolic link.')
            manifest = dict(format='redink-log-archive-v1', month=month, entries={})
            fd, name = tempfile.mkstemp(suffix='.tmp', prefix='.redink-', dir=dest)
            os.close(fd); temporary = Path(name)
            with ZipFile(temporary, 'w', ZIP_DEFLATED, allowZip64=True) as z:
                if archive.exists():
                    with ZipFile(archive) as old:
                        manifest = _verified_manifest(old, month)
                        for entry in manifest['entries']:
                            with old.open(entry) as src, z.open(entry, 'w', force_zip64=True) as dst:
                                shutil.copyfileobj(src, dst, 1024 * 1024)
                for p, hold, evidence in captures:
                    entry = p.name
                    if entry in manifest['entries'] and manifest['entries'][entry] != evidence:
                        entry = evidence['sha256'] + '/' + p.name
                    if entry not in manifest['entries']:
                        z.write(hold, entry)
                        manifest['entries'][entry] = evidence
                z.writestr(_MANIFEST, json.dumps(manifest, sort_keys=True))
            with ZipFile(temporary) as z:
                verified = _verified_manifest(z, month)
                if verified != manifest:
                    raise ValueError('Archive manifest changed during verification.')
            # Windows requires write access for the durable flush; r+b preserves the verified bytes.
            with temporary.open('r+b') as stream:
                os.fsync(stream.fileno())
            os.replace(temporary, archive); temporary = None
            # Reopen the actual committed archive before deleting any original.
            with ZipFile(archive) as z:
                _verified_manifest(z, month)
            for p, hold, evidence in captures:
                try:
                    with hold.open('rb') as stream:
                        if _digest(stream) != evidence['sha256']:
                            raise ValueError('Captured source changed; retain the source for another import.')
                    hold.unlink(); result['archived'] += 1
                    result['files'].append(dict(name=p.name, state='archived'))
                except (OSError, ValueError):
                    result['held'] += 1
                    result['files'].append(dict(name=p.name, state='retained', reason='Source changed or removal failed.'))
        except (ValueError, OSError, RuntimeError, BadZipFile, KeyError, UnicodeError, TypeError) as exc:
            result['held'] += len(captures)
            result['files'].append(dict(name=month, state='retained', reason=str(exc)))
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
            for p, hold, _ in captures:
                _restore(hold, p, result)
    return result


def archive_segments(store, before, destination, *, apply=False, today=None, allow_expired=False):
    from .reports import valid_month
    valid_month(before)
    today = today or date.today()
    if before > today.isoformat()[:7]:
        raise ValueError('A future boundary would include the current writable month.')
    dest = _destination(store, destination)
    if apply:
        with _archive_lock(dest):
            return _archive_segments(store, before, dest, apply=True, today=today, allow_expired=allow_expired)
    return _archive_segments(store, before, dest, apply=False, today=today, allow_expired=allow_expired)


def maintain_log_storage(store):
    """Run daily, before statistics retention; persist a bounded administrator status."""
    settings = store.settings
    today = store.clock().astimezone().date()
    policy = '|'.join(map(str, (today, settings.log_storage_archive_enabled,
        settings.log_storage_online_months, settings.log_storage_archive_months,
        settings.log_storage_archive_directory, settings.log_paths, store.log_retention_cutoff())))
    if store.metadata('log_archive_policy') == policy:
        return
    summary = dict(date=today.isoformat(), archived=0, expired_archives=0, held=0, legacy_files=0, details=[])
    if not settings.log_storage_archive_enabled or not settings.log_paths:
        summary['disabled'] = True
    else:
        try:
            dest = _destination(store, settings.log_storage_archive_directory or settings.data_dir / 'log-archives')
            with _archive_lock(dest):
                result = _archive_segments(store, month_boundary(settings.log_storage_online_months, today),
                    dest, apply=True, today=today, allow_expired=True)
                summary.update(archived=result['archived'], held=result['held'], legacy_files=result['legacy_files'])
                summary['details'] = [x for x in result['files'] if x['state'] != 'archived'][:20]
                expiry = month_boundary(settings.log_storage_archive_months, today)
                # Only our verified monthly bundles expire. Never delete arbitrary ZIPs.
                for path in dest.glob('redink-????-??.zip'):
                    match = _BUNDLE.fullmatch(path.name)
                    if not match or match[1] >= expiry or path.is_symlink():
                        continue
                    try:
                        with ZipFile(path) as z:
                            manifest = _verified_manifest(z, match[1])
                        if any(e['newest_date'] >= expiry + '-01' for e in manifest['entries'].values()):
                            continue
                        path.unlink(); summary['expired_archives'] += 1
                    except (OSError, ValueError, BadZipFile, KeyError, UnicodeError, TypeError):
                        summary['held'] += 1
                        if len(summary['details']) < 20:
                            summary['details'].append(dict(name=path.name, state='retained', reason='Archive expiry verification failed.'))
        except (ValueError, OSError, RuntimeError, BadZipFile) as exc:
            summary['held'] += 1
            summary['details'].append(dict(state='retained', reason=type(exc).__name__ + ': check archive location and permissions.'))
    store.save_metadata('log_archive_status', json.dumps(summary))
    store.save_metadata('log_archive_policy', policy)
    logger.log(logging.WARNING if summary['held'] else logging.INFO,
               'Log maintenance: archived=%s expired_archives=%s held=%s. Check Settings for details.',
               summary['archived'], summary['expired_archives'], summary['held'])
