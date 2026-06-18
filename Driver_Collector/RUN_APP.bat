@echo off
setlocal
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src;%PYTHONPATH%"

where python >nul 2>nul
if %errorlevel%==0 (
    python -m driver_collector_tool %*
    exit /b %errorlevel%
)

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 -m driver_collector_tool %*
    exit /b %errorlevel%
)

echo Python 3 was not found. Install Python 3.10 or newer.
exit /b 1
