@echo off
setlocal
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src;%PYTHONPATH%"

rem Default to gui when no arguments supplied
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
exit /b 1
