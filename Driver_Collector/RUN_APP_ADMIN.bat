@echo off
rem ─────────────────────────────────────────────────────────────────────────
rem  RUN_APP_ADMIN.bat  –  Launch Driver Collector GUI with Administrator rights
rem  No venv required. Uses system Python (python / py -3).
rem ─────────────────────────────────────────────────────────────────────────
setlocal

rem Check if already elevated
net session >nul 2>&1
if %errorlevel%==0 goto :run_elevated

rem Not elevated – re-launch this script via PowerShell runas
echo Requesting Administrator privileges...
powershell -NoProfile -Command ^
  "Start-Process cmd -ArgumentList '/c ""%~f0""' -Verb RunAs -WorkingDirectory '%~dp0'"
exit /b

:run_elevated
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src;%PYTHONPATH%"
set ARGS=%*
if "%ARGS%"=="" set ARGS=gui

where python >nul 2>nul
if %errorlevel%==0 (
    python -m driver_collector_tool %ARGS%
    exit /b %errorlevel%
)

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 -m driver_collector_tool %ARGS%
    exit /b %errorlevel%
)

echo Python 3 was not found. Install Python 3.10 or newer.
pause
exit /b 1
