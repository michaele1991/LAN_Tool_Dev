@echo off
setlocal
cd /d "%~dp0"

:: ── Validate or create venv ───────────────────────────────────────────────
set VENV_OK=0
if exist ".venv\Scripts\pythonw.exe" (
    .venv\Scripts\python.exe -c "import sys" >nul 2>&1
    if not errorlevel 1 set VENV_OK=1
)

if "%VENV_OK%"=="0" (
    echo Setting up virtual environment...
    if exist ".venv" rmdir /s /q .venv
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Python not found. Install Python 3.10+ from https://www.python.org
        pause
        exit /b 1
    )
    .venv\Scripts\python.exe -m pip install -q --upgrade pip --index-url https://pypi.org/simple/
    .venv\Scripts\python.exe -m pip install -q -r requirements.txt --index-url https://pypi.org/simple/
)

:: ── Launch app (no console window) ────────────────────────────────────────
start "" .venv\Scripts\pythonw.exe src\app.py
exit /b 0
