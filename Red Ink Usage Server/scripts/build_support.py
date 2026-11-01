"""Build helpers: pass only file paths and simple arguments across PowerShell's native boundary."""
import argparse
from pathlib import Path
import struct
import sys
import zipfile


def main(argv=None):
    cli = argparse.ArgumentParser()
    subs = cli.add_subparsers(dest='command', required=True)
    subs.add_parser('check-python')
    subs.add_parser('python-version')
    archive = subs.add_parser('archive')
    archive.add_argument('root', type=Path)
    archive.add_argument('target', type=Path)
    args = cli.parse_args(argv)
    if args.command == 'python-version':
        print(sys.version)
    elif args.command == 'check-python':
        print(sys.version)
        print('Architecture:', struct.calcsize('P') * 8)
        if sys.version_info < (3, 11):
            cli.exit(1, 'Python 3.11+ required\n')
    else:
        root, target = args.root.resolve(strict=True), args.target.resolve()
        if not root.is_dir() or target.is_relative_to(root):
            cli.error('Archive output must be outside the release directory.')
        with zipfile.ZipFile(target, 'x', zipfile.ZIP_DEFLATED) as output:
            for entry in sorted(root.rglob('*')):
                if entry.is_file():
                    output.write(entry, entry.relative_to(root.parent))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
