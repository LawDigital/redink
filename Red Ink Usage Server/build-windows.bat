@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-windows.ps1" %*
set "REDINK_EXIT=%ERRORLEVEL%"
if not "%REDINK_EXIT%"=="0" echo Build failed. See the error above.
exit /b %REDINK_EXIT%
