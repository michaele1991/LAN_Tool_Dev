@echo off
setlocal
cd /d "%~dp0"

:: ── Install dependencies if needed ────────────────────────────────────────
if not exist ".venv\Scripts\pythonw.exe" (
    echo Setting up virtual environment...
    python -m venv .venv
    .venv\Scripts\python.exe -m pip install -q --upgrade pip
    .venv\Scripts\python.exe -m pip install -q -r requirements.txt
)

:: ── Launch app (no console window) ────────────────────────────────────────
start "" .venv\Scripts\pythonw.exe src\app.py
exit /b 0
