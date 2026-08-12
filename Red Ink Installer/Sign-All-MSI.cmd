@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Test-PowerShellSyntax.ps1"
if errorlevel 1 (
  echo.
  echo INSTALLER POWERSHELL SYNTAX CHECK FAILED.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Sign-All-MSI.ps1"
if errorlevel 1 (
  echo.
  echo MSI SIGNING FAILED.
  pause
  exit /b 1
)
echo.
echo ALL RELEASE MSI FILES SIGNED AND VERIFIED.
pause
