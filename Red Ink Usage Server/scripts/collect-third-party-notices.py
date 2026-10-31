"""Copy exact installed license notices into each Windows release and enforce allowlist.

Run with the isolated build interpreter, which has all pinned dependencies.
Abort on unknown dependencies rather than silently shipping them.
"""
from __future__ import annotations
import importlib.metadata as metadata
import json
import pathlib
import shutil
import sys

ALLOWED = {
    'flask':'BSD-3-Clause', 'waitress':'ZPL-2.1', 'werkzeug':'BSD-3-Clause',
    'jinja2':'BSD-3-Clause', 'itsdangerous':'BSD-3-Clause', 'click':'BSD-3-Clause',
    'blinker':'MIT', 'markupsafe':'BSD-3-Clause', 'cheroot':'BSD-3-Clause',
    'cryptography':'Apache-2.0 OR BSD-3-Clause', 'more-itertools':'MIT',
    'jaraco.functools':'MIT', 'cffi':'MIT', 'pycparser':'BSD-3-Clause',
    'pyinstaller':'GPL-2.0-or-later WITH Bootloader-exception',
}


def main(out: pathlib.Path):
    out.mkdir(parents=True,exist_ok=True)
    from pathlib import Path
    base=Path(__file__).resolve().parents[1]
    requirements=[]
    for name in ('requirements.txt','requirements-build.txt'):
        requirements.extend(line.strip().split('==')[0].strip().lower() for line in (base/name).read_text().splitlines()
                            if '==' in line and not line.lstrip().startswith('#'))
    names=set(requirements)
    if names-set(ALLOWED):
        raise SystemExit('Unreviewed runtime/build dependency: '+', '.join(sorted(names-set(ALLOWED))))
    records=[]
    for name in sorted(names):
        dist=metadata.distribution(name)
        expected=ALLOWED[name]
        files=list(dist.files or ())
        candidates=[f for f in files if f.name.lower().startswith(('license','copying','notice')) and '.dist-info/' in str(f)]
        if not candidates:
            raise SystemExit('Installed package has no embedded license notice: '+name)
        files_out=[]
        package_dir=out/name
        package_dir.mkdir(exist_ok=True)
        for f in candidates:
            location=dist.locate_file(f)
            if location.is_file():
                target=package_dir/f.name
                shutil.copyfile(location,target)
                files_out.append(str(target.relative_to(out)))
        if not files_out:
            raise SystemExit('Could not read license notices for: '+name)
        records.append({'distribution':name,'version':dist.version,'license_reviewed_as':expected,'notice_files':files_out})
    (out/'manifest.json').write_text(json.dumps(records,indent=2)+'\n',encoding='utf-8')
    (out/'README.txt').write_text(
        'THIRD-PARTY SOFTWARE LICENCES AND NOTICES\n\n'
        'This application includes components provided by third parties.\n'
        'The component names, versions, applicable license texts and references\n'
        'are provided here and in manifest.json. Their respective licensors\n'
        'retain the rights and grant the permissions stated in those notices.\n\n'
        'PyInstaller bootloader: GPL-2.0-or-later WITH Bootloader Exception.\n'
        'https://pyinstaller.org/en/stable/license.html\n'
        'https://github.com/pyinstaller/pyinstaller\n', encoding='utf-8')
    print('Collected',len(records),'installed package license notices.')

if __name__=='__main__':
    if len(sys.argv)!=2: raise SystemExit('Pass destination directory for license notices.')
    main(pathlib.Path(sys.argv[1]).resolve())
