"""Deterministic cleanup of SQLite test fixtures across Windows file scanners.

Store.connect and backup operations close connections explicitly. On Windows,
anti-virus or other system services may briefly hold a SQLite file at the
moment a TemporaryDirectory is removed. Retry only this *transient* Windows
sharing violation, and fail loudly if a handle remains held. Never use
ignore_cleanup_errors (which could silently leave test data behind).
"""
from __future__ import annotations

import gc
import os
import tempfile
import time


class SqliteTemporaryDirectory(tempfile.TemporaryDirectory):
    def cleanup(self):
        for attempt in range(21):
            try:
                return super().cleanup()
            except PermissionError as exc:
                if os.name != 'nt' or getattr(exc, 'winerror', None) != 32 or attempt == 20:
                    raise
                gc.collect()
                time.sleep(min(0.03 * (attempt + 1), 0.3))
