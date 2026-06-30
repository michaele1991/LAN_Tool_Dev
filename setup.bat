@echo off
:: GBE NVM Builder — one-click setup
:: Creates a Python virtual environment and installs dependencies.
:: Run once after cloning.

echo.
echo  GBE NVM Builder — Environment Setup
echo  =====================================

:: Check Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  ERROR: Python not found. Install Python 3.10+ from https://python.org
    pause
    exit /b 1
)

echo  Creating virtual environment (.venv)...
python -m venv .venv

echo  Installing dependencies...
.venv\Scripts\python.exe -m pip install --quiet --upgrade pip
.venv\Scripts\python.exe -m pip install --quiet openpyxl

echo.
echo  Setup complete!
echo.
echo  To build a GBE NVM image, run:
echo    build.bat
echo.
pause
