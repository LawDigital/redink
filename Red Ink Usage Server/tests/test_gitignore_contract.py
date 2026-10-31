"""Git integration regression: source fixtures tracked, secrets ignored.

No changes to Red Ink.sln or the parent .gitignore are necessary.
"""
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

SERVER = Path(__file__).resolve().parents[1]


@unittest.skipUnless(shutil.which('git'), 'Git is not installed on this test host')
class GitIgnoreContractTests(unittest.TestCase):
    def test_server_inside_visual_studio_repository(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            # The root gitignore is taken verbatim from the uploaded VB solution.
            (root / '.gitignore').write_text(
                'clickonce/\npackages/\nbin/\nobj/\n*.user\n*.suo\n*.vs/\n\n/CONTRIBUTING.md\n',
                encoding='utf-8')
            server = root / 'Red Ink Usage Server'
            server.mkdir()
            (server / '.gitignore').write_bytes((SERVER / '.gitignore').read_bytes())
            tracked = [
                'redink_server/web.py', 'redink_server/static/app.js',
                'redink_server/static/red-ink-logo.png',
                'tests/data/counter-vb-vectors-v2.json',
                'config/server.example.toml', 'build/red-ink.ico',
                'dist/README.txt', 'README.md',
                'docs/ADMINISTRATOR_GUIDE.md', 'docs/USER_GUIDE.md',
                'docs/THIRD_PARTY_LICENSES.md',
            ]
            ignored = [
                '__pycache__/test.pyc', '.venv/pyvenv.cfg', '.venv-build/pyvenv.cfg',
                'data/redink.sqlite3', 'secrets.json', 'server.toml',
                'tls/key.pem', 'certificates/cert.pfx',
                'dist/windows-onedir/RedInkUsageServer.exe',
                'build/pyinstaller-onedir/warn.txt',
                'logs/redink-0123456789ABCDEF-WD.log',
                'docs/QA.md', 'docs/QA_1_3.md', 'docs/validation/worker-audit.txt',
                'docs/addin-changes.diff', 'docs/WINDOWS.md',
                'SHA256SUMS.txt', 'build-info.json', 'build-dependencies.txt',
            ]
            for rel in tracked + ignored:
                target = server / rel
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_text('test', encoding='utf-8')
            subprocess.run(['git', 'init', '-q'], cwd=root, check=True,
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            subprocess.run(['git', 'add', '-A', '--', 'Red Ink Usage Server'],
                           cwd=root, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            found = subprocess.run(['git', 'diff', '--cached', '--name-only', '-z'],
                                   cwd=root, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            staged = {p.decode('utf-8') for p in found.stdout.split(b'\x00') if p}
            for rel in tracked:
                self.assertIn(f'Red Ink Usage Server/{rel}', staged)
            for rel in ignored:
                self.assertNotIn(f'Red Ink Usage Server/{rel}', staged)


if __name__ == '__main__':
    unittest.main()
