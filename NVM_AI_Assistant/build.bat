@echo off
:: GBE NVM Image Builder — 5-question wizard
setlocal
cd /d "%~dp0"

set VENV=.venv\Scripts\python.exe
set WIZARD=src\wizard.py

:: If venv doesn't exist, create it from requirements.txt
if not exist "%VENV%" (
    echo  Setting up virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo  ERROR: Python not found. Install Python 3.10+ from https://www.python.org
        pause
        exit /b 1
    )
    .venv\Scripts\python.exe -m pip install --quiet -r requirements.txt
)

echo.
"%VENV%" "%WIZARD%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  Build exited with an error. See messages above.
    pause
)
