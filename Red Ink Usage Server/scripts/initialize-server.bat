@echo off
setlocal
set "REDINK_ROOT=%~dp0"
if not exist "%REDINK_ROOT%RedInkUsageServer.exe" set "REDINK_ROOT=%~dp0..\"
if not exist "%REDINK_ROOT%RedInkUsageServer.exe" (
  echo RedInkUsageServer.exe was not found. Run this script from the Windows distribution.
  exit /b 1
)
"%REDINK_ROOT%RedInkUsageServer.exe" init --data-dir "%ProgramData%\RedInkUsageServer" %*
set "REDINK_EXIT=%ERRORLEVEL%"
if "%REDINK_EXIT%"=="0" echo Save the generated administrator password and tokens securely before closing this window.
pause
exit /b %REDINK_EXIT%
