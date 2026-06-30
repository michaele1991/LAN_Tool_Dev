@echo off
:: GBE NVM Builder — launch the 5-question build wizard
setlocal

set VENV=.venv\Scripts\python.exe
set WIZARD=GBE_Builder\wizard.py

:: If venv doesn't exist, run setup first
if not exist "%VENV%" (
    echo  Virtual environment not found — running setup first...
    call setup.bat
)

echo.
"%VENV%" "%WIZARD%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  Build exited with an error. See messages above.
    pause
)
