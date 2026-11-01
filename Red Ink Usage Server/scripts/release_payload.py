"""Stage and verify the public Windows runtime package; never copy development trees.

Only this allowlist defines distributable project files. PyInstaller's binary
output and exact third-party notices are populated by the Windows build.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

PUBLIC_DOCS = (
    'ADMINISTRATOR_GUIDE.md',
    'USER_GUIDE.md',
)
PUBLIC_SCRIPTS = (
    'user-autostart.ps1',
    'install-startup-task.ps1',
    'uninstall-startup-task.ps1',
    'run-server-with-tls-password.ps1',
)
PUBLIC_LAUNCHERS = ('run-server.bat', 'initialize-server.bat')
ALLOWED_RELEASE_ROOT = {
    'RedInkUsageServer.exe', '_internal', 'README.txt', 'docs',
    'scripts', 'third-party-licenses', *PUBLIC_LAUNCHERS,
}
NEVER_SHIP = {
    'build-info.json', 'build-dependencies.txt', 'requirements.txt',
    'requirements-build.txt', 'server.toml', 'secrets.json',
    'run_server.py', 'RedInkUsageServer.spec', 'build-windows.ps1',
    'build-windows.bat',
}


def stage_public_files(source: Path, release: Path) -> None:
    """Add the minimal operational files to an already frozen onedir/onefile tree."""
    if not (release / 'RedInkUsageServer.exe').is_file():
        raise ValueError('Signed or unsigned frozen executable is missing')
    selected = [('packaging/README.txt', 'README.txt')]
    selected.extend((f'docs/{name}', f'docs/{name}') for name in PUBLIC_DOCS)
    selected.extend((f'scripts/{name}', f'scripts/{name}') for name in PUBLIC_SCRIPTS)
    selected.extend((f'scripts/{name}', name) for name in PUBLIC_LAUNCHERS)
    for origin_name, destination_name in selected:
        origin = source / origin_name
        if not origin.is_file() or origin.is_symlink():
            raise ValueError(f'Required public file is missing or linked: {origin_name}')
        destination = release / destination_name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(origin, destination)


def verify_public_release(release: Path) -> None:
    """Fail closed if unwanted source, test files or credentials entered the ZIP."""
    if not (release / 'RedInkUsageServer.exe').is_file():
        raise ValueError('Frozen executable missing')
    roots = {item.name for item in release.iterdir()}
    unexpected_roots = sorted(roots - ALLOWED_RELEASE_ROOT)
    if unexpected_roots:
        raise ValueError('Unexpected release-root entries: ' + ', '.join(unexpected_roots))
    expected = {'README.txt', *PUBLIC_LAUNCHERS}
    expected.update('docs/' + d for d in PUBLIC_DOCS)
    expected.update('scripts/' + s for s in PUBLIC_SCRIPTS)
    for relative in sorted(expected):
        if not (release / relative).is_file():
            raise ValueError('Public runtime file missing: ' + relative)
    if not (release / 'third-party-licenses' / 'README.txt').is_file():
        raise ValueError('Third-party license notice directory missing')
    if not (release / 'third-party-licenses' / 'manifest.json').is_file():
        raise ValueError('Third-party license manifest missing')
    # Strictly review the own docs/scripts; exact upstream notices and frozen
    # interpreter assets are allowed in their dedicated, separately audited trees.
    for folder, names in (('docs', set(PUBLIC_DOCS)),
                          ('scripts', set(PUBLIC_SCRIPTS))):
        actual = {p.name for p in (release / folder).iterdir()}
        if actual != names:
            raise ValueError(f'Unexpected {folder} files: {sorted(actual - names)}; missing: {sorted(names - actual)}')
    for path in release.rglob('*'):
        if path.is_symlink():
            raise ValueError('Linked file in release: ' + path.relative_to(release).as_posix())
        if not path.is_file():
            continue
        rel = path.relative_to(release)
        name = path.name.lower()
        if rel.parts[0] == 'third-party-licenses':
            continue  # Genuine, unmodified upstream notices have independent names.
        if rel.parts[0] == '_internal' and any(
            folder.lower() in ('tests', 'test', 'examples', 'samples',
                               'docs', 'documentation', '__pycache__')
            for folder in rel.parts[1:-1]
        ):
            raise ValueError('Development directory in runtime: ' + rel.as_posix())
        if name in {entry.lower() for entry in NEVER_SHIP} or name.endswith((
            '.py', '.pyw', '.pyi', '.spec', '.diff', '.patch', '.pdb', '.obj',
            '.c', '.h', '.cs', '.vb', '.sqlite3', '.sqlite3-wal',
            '.sqlite3-shm', '.pfx', '.p12', '.pem', '.key', '.log', '.log.1',
        )):
            raise ValueError('Prohibited file in release: ' + rel.as_posix())
        if rel.parts[0] not in ('_internal', 'docs', 'scripts') and name.endswith(('.json', '.md')):
            raise ValueError('Unexpected metadata/document: ' + rel.as_posix())


def main(argv: list[str]) -> int:
    if len(argv) != 3 or argv[0] not in ('stage', 'verify'):
        raise SystemExit('usage: release_payload.py stage SOURCE RELEASE | verify SOURCE RELEASE')
    operation, source_name, release_name = argv
    source, release = Path(source_name).resolve(), Path(release_name).resolve()
    if operation == 'stage':
        stage_public_files(source, release)
    else:
        verify_public_release(release)
    print(f'Public Windows release {operation}: OK')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
