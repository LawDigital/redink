"""Create the cross-platform Python release using only the standard library."""
from __future__ import annotations

import re
import runpy
import stat
import sys
import tempfile
import zipfile
from pathlib import Path


PRODUCT = 'RedInkUsageServer'
REQUIRED_FILES = (
    'run_server.py',
    'requirements.txt',
    'redink_server/__init__.py',
    'redink_server/cli.py',
    'redink_server/web.py',
    'redink_server/feature_taxonomy.default.json',
    'redink_server/static/app.css',
    'redink_server/static/app.js',
    'redink_server/static/print.js',
    'redink_server/static/red-ink-logo.png',
    'redink_server/static/report.css',
    'redink_server/templates/index.html',
    'redink_server/templates/login.html',
    'redink_server/templates/report.html',
    'config/server.example.toml',
    'INSTALL_PYTHON.md',
    'docs/ADMINISTRATOR_GUIDE.md',
    'docs/USER_GUIDE.md',
    'docs/THIRD_PARTY_LICENSES.md',
    'scripts/user-autostart.ps1',
)


def checked_file(root: Path, relative: str) -> Path:
    path = root
    for part in Path(relative).parts:
        path /= part
        info = path.lstat()
        if stat.S_ISLNK(info.st_mode) or (
            getattr(info, 'st_file_attributes', 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT
        ):
            raise ValueError(f'Linked release input: {relative}')
    if not path.is_file():
        raise ValueError(f'Required release input is not a regular file: {relative}')
    return path


def build_release(root: Path) -> Path:
    # Runtime modules are discovered only in the flat application package. All
    # other resources are explicitly named, including the sample configuration.
    files = {name: checked_file(root, name) for name in REQUIRED_FILES}
    for module in (root / 'redink_server').glob('*.py'):
        relative = module.relative_to(root).as_posix()
        files[relative] = checked_file(root, relative)
    version = runpy.run_path(str(files['redink_server/__init__.py'])).get('__version__')
    if not isinstance(version, str) or not re.fullmatch(r'[0-9]+(?:\.[0-9]+)+(?:[-+][A-Za-z0-9.]+)?', version):
        raise ValueError('Invalid package version in redink_server/__init__.py')

    output = root / 'dist' / 'python'
    output.mkdir(parents=True, exist_ok=True)
    target = output / f'{PRODUCT}-{version}-python.zip'
    with tempfile.TemporaryDirectory(prefix='.release-', dir=output) as temporary:
        archive = Path(temporary) / target.name
        with zipfile.ZipFile(archive, 'x', zipfile.ZIP_DEFLATED) as zipped:
            for relative, path in sorted(files.items()):
                entry = 'README.md' if relative == 'INSTALL_PYTHON.md' else relative
                zipped.write(path, f'{PRODUCT}/{entry}')
        with zipfile.ZipFile(archive) as zipped:
            if zipped.testzip() is not None or len(zipped.infolist()) != len(files):
                raise ValueError('ZIP integrity check failed.')
        archive.replace(target)
    return target


if __name__ == '__main__':
    if sys.version_info < (3, 11):
        sys.exit('Python 3.11 or later is required.')
    try:
        print(f'Python distribution archive: {build_release(Path(__file__).resolve().parents[1])}')
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        sys.exit(f'Python release failed: {error}')
