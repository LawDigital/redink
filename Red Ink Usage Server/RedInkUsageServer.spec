# Run on Windows using build-windows.ps1. Mutable data is never bundled here.
import os
from pathlib import Path

project = Path(SPECPATH).resolve()
onefile = os.environ.get("REDINK_BUILD_MODE", "onedir") == "onefile"

analysis = Analysis(
    [str(project / "run_server.py")],
    pathex=[str(project)],
    binaries=[],
    datas=[
        (str(project / "redink_server" / "templates"), "redink_server/templates"),
        (str(project / "redink_server" / "static"), "redink_server/static"),
        (str(project / "redink_server" / "feature_taxonomy.default.json"), "redink_server"),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
archive = PYZ(analysis.pure)
options = dict(
    name="RedInkUsageServer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    icon=str(project / "build" / "red-ink.ico"),
    version=str(project / "build" / "version_info.txt"),
)
if onefile:
    executable = EXE(archive, analysis.scripts, analysis.binaries, analysis.datas, [], **options)
else:
    executable = EXE(archive, analysis.scripts, [], exclude_binaries=True, **options)
    distribution = COLLECT(executable, analysis.binaries, analysis.datas,
                           strip=False, upx=False, name="RedInkUsageServer")
