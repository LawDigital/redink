@echo off
setlocal
set "PS_SCRIPT=%~dp0create-source-zip-on-desktop.ps1"
if not exist "%PS_SCRIPT%" (
  echo ERROR: Missing companion PowerShell script: "%PS_SCRIPT%"
  exit /b 2
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
if errorlevel 1 (
  echo ERROR: Source archive was not created.
  exit /b 1
)
endlocal
