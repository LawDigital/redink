@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Test-PowerShellSyntax.ps1"
if errorlevel 1 (
  echo.
  echo INSTALLER POWERSHELL SYNTAX CHECK FAILED.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Enable-VdprojCommandLineBuild.ps1"
if errorlevel 1 (
  echo.
  echo VDPROJ WORKAROUND FAILED.
  pause
  exit /b 1
)
echo.
pause
