@echo off
setlocal
where py.exe >nul 2>&1
if errorlevel 1 goto use_python
py.exe -3 "%~dp0scripts\build-python-release.py"
goto finished

:use_python
python.exe "%~dp0scripts\build-python-release.py"

:finished
set "REDINK_EXIT=%ERRORLEVEL%"
if not "%REDINK_EXIT%"=="0" echo Python package build failed. Python 3.11 or later is required.
exit /b %REDINK_EXIT%
