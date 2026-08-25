@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Test-PowerShellSyntax.ps1"
if errorlevel 1 (
  echo.
  echo INSTALLER POWERSHELL SYNTAX CHECK FAILED.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Test-InstallerEnvironment.ps1"
if errorlevel 1 (
  echo.
  echo INSTALLER SETUP CHECK FAILED.
  pause
  exit /b 1
)
echo.
echo INSTALLER SETUP CHECK COMPLETED.
pause
