"""Current-user Windows logon startup control; no administrator privileges needed.

The Task Scheduler task runs the same executable/script and data directory as
this server. No mutable web-form path or user-supplied command is executed.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

TASK_NAME = 'RedInkUsageServer-CurrentUser'


class StartupError(ValueError):
    pass


def _script() -> Path:
    if getattr(sys, 'frozen', False):
        script = Path(sys.executable).resolve().parent / 'scripts' / 'user-autostart.ps1'
    else:
        script = Path(__file__).resolve().parents[1] / 'scripts' / 'user-autostart.ps1'
    if not script.is_file():
        raise StartupError('Autostart helper script is missing from this installation.')
    return script


def _execution() -> tuple[str, str]:
    if getattr(sys, 'frozen', False):
        return str(Path(sys.executable).resolve()), ''
    return str(Path(sys.executable).resolve()), str(Path(__file__).resolve().parents[1] / 'run_server.py')


def manage_autostart(data_dir: Path, action: str) -> dict:
    """Manage only the fixed current-user task. Does not start/stop server."""
    if action not in ('status', 'enable', 'disable'):
        raise StartupError('Unsupported autostart action.')
    if os.name != 'nt':
        raise StartupError('Windows autostart is only available on Windows.')
    exe, script_path = _execution()
    command = ['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
               '-File', str(_script()), '-Action', action, '-Executable', exe,
               '-DataDir', str(data_dir.resolve())]
    if script_path:
        command.extend(['-ScriptPath', script_path])
    try:
        completed = subprocess.run(command, capture_output=True, text=True, timeout=25, check=False,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise StartupError('Windows Task Scheduler could not be contacted.') from exc
    if completed.returncode:
        # Do not return unchecked PowerShell stdout: it might include paths or secrets.
        raise StartupError('Windows Task Scheduler rejected this operation. Check Task Scheduler permissions and the server log.')
    return {'enabled': completed.stdout.strip() == 'ENABLED', 'mode': 'current-user-logon',
            'headless': True, 'requires_windows_logon': True}
