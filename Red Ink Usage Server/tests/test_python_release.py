"""Release tests for the Python distribution, independent of runtime dependencies."""
import importlib.util
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location('build_python_release', ROOT / 'scripts/build-python-release.py')
release = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(release)


class PythonReleaseTests(unittest.TestCase):
    def test_release_filters_local_state_and_preserves_previous_zip_on_missing_asset(self):
        with tempfile.TemporaryDirectory() as temporary:
            project = Path(temporary)
            for name in release.REQUIRED_FILES:
                destination = project / name
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(ROOT / name, destination)
            (project / 'redink_server' / 'additional.py').write_text('VALUE = 1\n', encoding='utf-8')
            for name in ('secrets.json', 'redink_server/secrets.json', 'redink_server/static/server.log'):
                (project / name).write_text('private', encoding='utf-8')
            target = release.build_release(project)
            self.assertEqual(target.parent, project / 'dist' / 'python')
            self.assertRegex(target.name, r'^RedInkUsageServer-\d.*-python\.zip$')
            self.assertEqual(release.build_release(project), target)
            with zipfile.ZipFile(target) as archive:
                self.assertIsNone(archive.testzip())
                members = set(archive.namelist())
                expected = set(release.REQUIRED_FILES) - {'INSTALL_PYTHON.md'}
                expected.update({'README.md', 'redink_server/additional.py'})
                self.assertEqual(members, {'RedInkUsageServer/' + entry for entry in expected})
            previous = target.read_bytes()
            (project / 'redink_server/templates/login.html').unlink()
            with self.assertRaises(FileNotFoundError):
                release.build_release(project)
            self.assertEqual(target.read_bytes(), previous)
