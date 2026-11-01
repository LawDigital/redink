"""Regression contract: public Windows ZIP contains runtime and public docs only."""
import importlib.util
import tempfile
import zipfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('public_release', ROOT / 'scripts' / 'release_payload.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class PublicReleaseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.release = Path(self.tmp.name) / 'RedInkUsageServer'
        self.release.mkdir()
        (self.release / 'RedInkUsageServer.exe').write_bytes(b'fake-executable')
        (self.release / '_internal').mkdir()
        (self.release / '_internal' / 'python312.dll').write_bytes(b'fake-library')
        notices = self.release / 'third-party-licenses'
        notices.mkdir()
        (notices / 'README.txt').write_text('Notice', encoding='utf-8')
        (notices / 'manifest.json').write_text('[]', encoding='utf-8')

    def tearDown(self):
        self.tmp.cleanup()

    def test_public_package_is_exact_allowlist(self):
        module.stage_public_files(ROOT, self.release)
        module.verify_public_release(self.release)
        self.assertTrue((self.release / 'docs' / 'ADMINISTRATOR_GUIDE.md').exists())
        self.assertTrue((self.release / 'docs' / 'USER_GUIDE.md').exists())
        self.assertTrue((self.release / 'scripts' / 'user-autostart.ps1').exists())
        self.assertFalse((self.release / 'docs' / 'QA.md').exists())
        self.assertFalse((self.release / 'build-info.json').exists())
        self.assertFalse((self.release / 'requirements.txt').exists())
        self.assertFalse((self.release / 'scripts' / 'collect-third-party-notices.py').exists())

    def test_development_root_directory_rejected(self):
        module.stage_public_files(ROOT, self.release)
        (self.release / 'tests').mkdir()
        with self.assertRaisesRegex(ValueError, 'Unexpected release-root'):
            module.verify_public_release(self.release)

    def test_source_code_leak_rejected_in_internal_tree(self):
        module.stage_public_files(ROOT, self.release)
        (self.release / '_internal' / 'run_server.py').write_text('print(1)')
        with self.assertRaisesRegex(ValueError, 'Prohibited file'):
            module.verify_public_release(self.release)

    def test_distribution_zip_has_only_public_content(self):
        module.stage_public_files(ROOT, self.release)
        module.verify_public_release(self.release)
        archive = Path(self.tmp.name) / 'distribution.zip'
        with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for source in sorted(self.release.rglob('*')):
                if source.is_file():
                    zip_out.write(source, source.relative_to(self.release.parent))
        with zipfile.ZipFile(archive) as zip_in:
            self.assertIsNone(zip_in.testzip())
            names = set(zip_in.namelist())
            self.assertIn('RedInkUsageServer/RedInkUsageServer.exe', names)
            self.assertIn('RedInkUsageServer/docs/USER_GUIDE.md', names)
            self.assertIn('RedInkUsageServer/docs/ADMINISTRATOR_GUIDE.md', names)
            self.assertNotIn('RedInkUsageServer/tests/test_public_release.py', names)
            self.assertNotIn('RedInkUsageServer/build-info.json', names)
            self.assertNotIn('RedInkUsageServer/README.md', names)

    def test_development_directory_in_internal_tree_rejected(self):
        module.stage_public_files(ROOT, self.release)
        directory = self.release / '_internal' / 'tests'
        directory.mkdir()
        (directory / 'test_suite.txt').write_text('test')
        with self.assertRaisesRegex(ValueError, 'Development directory'):
            module.verify_public_release(self.release)

    def test_developer_document_rejected_in_docs(self):
        module.stage_public_files(ROOT, self.release)
        (self.release / 'docs' / 'QA_1_3.md').write_text('internal')
        with self.assertRaisesRegex(ValueError, 'Unexpected docs files'):
            module.verify_public_release(self.release)

    def test_sensitive_data_file_rejected(self):
        module.stage_public_files(ROOT, self.release)
        (self.release / '_internal' / 'secrets.json').write_text('secret')
        with self.assertRaisesRegex(ValueError, 'Prohibited file'):
            module.verify_public_release(self.release)

    def test_missing_licence_notices_fails(self):
        module.stage_public_files(ROOT, self.release)
        (self.release / 'third-party-licenses' / 'manifest.json').unlink()
        with self.assertRaisesRegex(ValueError, 'license manifest missing'):
            module.verify_public_release(self.release)

    def test_ensure_no_recursively_copied_folders(self):
        source = (ROOT / 'build-windows.ps1').read_text(encoding='utf-8')
        self.assertIn("'scripts/release_payload.py', 'stage'", source)
        self.assertIn("'scripts/release_payload.py', 'verify'", source)
        self.assertNotIn("foreach ($directory in @('docs', 'config', 'examples', 'scripts'))", source)
        self.assertNotIn("'build-dependencies.txt'", source)
        self.assertNotIn("'build-info.json'", source)
        self.assertNotIn("'requirements-build.txt'))", source)
        self.assertIn('Invoke-Checked $Tool', source)

    def test_public_documentation_has_no_dangling_references(self):
        guide = (ROOT / 'docs' / 'ADMINISTRATOR_GUIDE.md').read_text(encoding='utf-8')
        for reference in ('docs/TLS.md', 'docs/WINDOWS.md', 'docs/QA.md', 'docs/DATABASE_MIGRATIONS.md', 'docs/ADDIN_CONFIGURATION.md'):
            self.assertNotIn(reference, guide)


if __name__ == '__main__':
    unittest.main()
