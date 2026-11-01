"""Conservative, opt-in archival of closed monthly SharedLogger streams.

The source is never modified by the normal scanner. On explicit apply, move
one closed segment out of the scanner namespace before checking it. The ZIP
is verified before the original is removed. A failed or interrupted archive
keeps the original or a recoverable hold file, never overwrites a new writer.
"""
from __future__ import annotations

from datetime import date
from hashlib import sha256
import os
from pathlib import Path
import re
import tempfile
from uuid import uuid4
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from .importers import safe_iter_files
from .log_parser import LogFraming, log_host_from_name, logical_log_name

_SEGMENT = re.compile(r'redink-[0-9a-f]{16}-(\d{4})(0[1-9]|1[0-2])-(WD|XL|EX|OL|UK)\.log',re.I)


def _digest(stream):
    h=sha256()
    for block in iter(lambda:stream.read(1024*1024),b''):
        h.update(block)
    return h.hexdigest()


def month_boundary(keep_months, today=None):
    """Keep current month and the previous N-1 complete calendar months."""
    if not 1 <= keep_months <= 1200:
        raise ValueError('keep-months must be between 1 and 1200.')
    now = today or date.today()
    year, month = divmod(now.year * 12 + now.month - keep_months, 12)
    return f'{year:04}-{month+1:02}'


def archive_segments(store, before, destination, *, apply=False, today=None):
    from .reports import valid_month
    valid_month(before)
    if before > (today or date.today()).isoformat()[:7]:
        raise ValueError('A future boundary would include the current writable month.')
    dest=Path(destination).resolve()
    roots=store.settings.log_paths
    protected=(*roots,*store.settings.counter_paths,store.settings.config_directory)
    if any(dest.is_relative_to(Path(root).resolve()) for root in protected):
        raise ValueError('Use an archive directory outside publishing and import roots.')
    result={'mode':'apply' if apply else 'preview','eligible':0,'archived':0,'held':0,'files':[]}
    visited=set()
    for root in roots:
        for p in safe_iter_files(root,recursive=store.settings.recursive,suffixes=('.log',)):
            match=_SEGMENT.fullmatch(p.name)
            if p in visited or not match or f'{match[1]}-{match[2]}' >= before:
                continue
            visited.add(p)
            hold=None; tmp=None
            try:
                status=store.file_status('log',p); cursor=store.log_scan_cursor(p); st=p.stat()
                if (not status or not cursor or status['errors'] or status['size']!=st.st_size or
                    status['mtime_ns']!=st.st_mtime_ns or cursor['offset_bytes']!=st.st_size or
                    not cursor['head_signature'].startswith('fr2:') or not status['records']):
                    raise ValueError('No current, complete, error-free import.')
                if apply:
                    dest.mkdir(mode=0o700,parents=True,exist_ok=True)
                    # Same volume, unique name. Capture first; do not later unlink a replacement at p.
                    candidate=p.with_name(p.name+'.archive-hold-'+uuid4().hex)
                    p.rename(candidate); hold=candidate
                source=hold or p
                # Read the captured file under its logical name, not the temporary hold suffix.
                from .importers import DEFAULT_MAX_LINE_BYTES
                framing=LogFraming(); seen=0; original=sha256(); length=0
                with source.open('rb') as stream, store.connect() as con:
                    con.execute('BEGIN')
                    for number,raw in enumerate(iter(lambda:stream.readline(DEFAULT_MAX_LINE_BYTES+1),b''),1):
                        if len(raw)>DEFAULT_MAX_LINE_BYTES or not raw.endswith(b'\n'):
                            raise ValueError('Incomplete or oversized record.')
                        length+=len(raw); original.update(raw)
                        text=raw.decode('utf-8-sig' if number==1 else 'utf-8')
                        fp=sha256(text.lstrip('\ufeff').rstrip('\r\n').encode()).hexdigest()
                        row=framing.line(text,log_host_from_name(p),number,fp)
                        if row.error: raise ValueError('Record did not pass framing checks.')
                        if row.event:
                            key=sha256(f'{logical_log_name(p)}|{row.event["host_id"]}|{fp}|{number}'.encode()).hexdigest()
                            if not con.execute('SELECT 1 FROM log_events WHERE record_id=?',(key,)).fetchone():
                                raise ValueError('Not every record is present in SQLite (possibly removed by retention).')
                            seen+=1
                    if length!=st.st_size or seen!=status['records']: raise ValueError('Source no longer matches the import.')
                result['eligible']+=1
                if apply:
                    archive=dest/(p.stem+'-'+original.hexdigest()+'.zip')
                    if not archive.exists():
                        fd, name=tempfile.mkstemp(suffix='.tmp',dir=dest);os.close(fd);tmp=Path(name)
                        with ZipFile(tmp,'w',ZIP_DEFLATED) as z: z.write(source,p.name)
                        with ZipFile(tmp) as z, z.open(p.name) as entry:
                            if _digest(entry)!=original.hexdigest(): raise ValueError('Archive verification failed.')
                        os.replace(tmp,archive);tmp=None
                    with ZipFile(archive) as z, z.open(p.name) as entry,source.open('rb') as current:
                        if _digest(entry)!=original.hexdigest() or _digest(current)!=original.hexdigest():
                            raise ValueError('Source or archive changed; keeping the original.')
                    source.unlink(); hold=None; result['archived']+=1
                result['files'].append({'name':p.name,'state':'archived' if apply else 'eligible'})
            except (ValueError, OSError, RuntimeError, BadZipFile, KeyError) as exc:
                result['held']+=1
                result['files'].append({'name':p.name,'state':'retained','reason':str(exc)})
            finally:
                if tmp is not None: tmp.unlink(missing_ok=True)
                if hold is not None:
                    # Never replace a stream recreated by an offline writer. No-clobber hard link;
                    # if unsupported, keep the hold file and report its recoverable location.
                    try:
                        os.link(hold,p);hold.unlink()
                    except OSError:
                        result['files'].append({'name':hold.name,'state':'manual recovery required'})
    return result
