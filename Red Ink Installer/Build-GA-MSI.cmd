@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Test-PowerShellSyntax.ps1"
if errorlevel 1 (
  echo.
  echo INSTALLER POWERSHELL SYNTAX CHECK FAILED.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Build-Channel.ps1" -Channel GA
if errorlevel 1 (
  echo.
  echo GA MSI BUILD FAILED.
  pause
  exit /b 1
)
echo.
echo GA MSI BUILD COMPLETED.
pause
