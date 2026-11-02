"""Static non-Windows contracts for opt-in device login and source archive."""
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]


class SupportArchiveAndBrandContractTests(unittest.TestCase):
    def test_source_archive_is_companion_batch_and_excludes_production_secrets(self):
        wrapper = (ROOT / 'scripts/create-source-zip-on-desktop.bat').read_text(encoding='utf-8')
        script = (ROOT / 'scripts/create-source-zip-on-desktop.ps1').read_text(encoding='utf-8')
        self.assertIn('create-source-zip-on-desktop.ps1', wrapper)
        self.assertIn('DesktopDirectory', script)
        self.assertIn('CreateEntryFromFile', script)
        self.assertIn('SOURCE_ARCHIVE_INFO.txt', script)
        self.assertIn("'redink_server'", script)
        self.assertIn("'tests'", script)
        for name in ('server.toml','secrets.json','.pfx','.pem','.sqlite3','config-files',
                     'backups',"'logs'",'dist'):
            self.assertIn(name, script)
        self.assertIn('redink_server\\web.py', script)
        self.assertIn('New-Object System.IO.Compression.ZipArchive', script)

    def test_remember_is_opt_in_and_logo_red_is_canonical(self):
        login = (ROOT / 'redink_server/templates/login.html').read_text(encoding='utf-8')
        css = (ROOT / 'redink_server/static/app.css').read_text(encoding='utf-8')
        self.assertIn('name="remember_device"', login)
        self.assertNotIn('name="remember_device" value="on" checked', login)
        self.assertIn('--brand-red:#b91312;', css)
        self.assertIn('accent-color:var(--brand-red)', css)
        # UI brand and data visualization colours have separate contracts.
        self.assertIn('--accent:var(--brand-red)', css)
        self.assertIn('--accent2:#ff5a36;', css)
        self.assertIn('--err:#ff5a5a;', css)
        self.assertIn('.version-addin-title{color:var(--brand-red)', css)
        self.assertIn('.legend-stroke.agent{background:var(--accent2)', css)
        self.assertIn('a{color:#ff7770;', css)
        self.assertIn('.source-kind{display:block;color:#f7a39c', css)
        self.assertNotIn('#c32137',css.lower())


if __name__ == '__main__':
    unittest.main()
