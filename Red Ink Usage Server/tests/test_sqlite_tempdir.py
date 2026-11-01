"""Bounded Windows-only cleanup fallback must never conceal real failures."""
import os
import tempfile
import unittest
from unittest.mock import patch
from sqlite_tempdir import SqliteTemporaryDirectory


class SqliteCleanupTests(unittest.TestCase):
    def test_normal_cleanup_removes_directory(self):
        t = SqliteTemporaryDirectory()
        path = t.name
        t.cleanup()
        self.assertFalse(os.path.exists(path))

    def test_nonsharing_permission_error_not_swallowed(self):
        t = SqliteTemporaryDirectory()
        error = PermissionError('not a sharing violation')
        with patch.object(tempfile.TemporaryDirectory, 'cleanup', side_effect=error):
            with self.assertRaises(PermissionError):
                t.cleanup()
        t.cleanup()

    def test_transient_windows_sharing_violation_retries_and_does_not_hide_it(self):
        import sqlite_tempdir
        t = SqliteTemporaryDirectory()
        failure = PermissionError('sharing violation')
        failure.winerror = 32
        with patch.object(sqlite_tempdir.os, 'name', 'nt'):
            with patch.object(tempfile.TemporaryDirectory, 'cleanup', side_effect=[failure, None]) as base:
                with patch.object(sqlite_tempdir.time, 'sleep'):
                    t.cleanup()
                self.assertEqual(base.call_count, 2)
        t.cleanup()

    def test_permanent_windows_sharing_violation_is_reported(self):
        import sqlite_tempdir
        t = SqliteTemporaryDirectory()
        failure = PermissionError('sharing violation')
        failure.winerror = 32
        with patch.object(sqlite_tempdir.os, 'name', 'nt'):
            with patch.object(tempfile.TemporaryDirectory, 'cleanup', side_effect=failure) as base:
                with patch.object(sqlite_tempdir.time, 'sleep'):
                    with self.assertRaises(PermissionError):
                        t.cleanup()
                self.assertEqual(base.call_count, 21)
        t.cleanup()

    def test_no_ignore_errors_flag_in_critical_cleanup(self):
        t = SqliteTemporaryDirectory()
        self.assertFalse(t._ignore_cleanup_errors)
        t.cleanup()
