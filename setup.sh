#!/usr/bin/env bash
# GBE NVM Builder — one-click setup (Linux / macOS)
# Creates a Python virtual environment and installs dependencies.
# Run once after cloning:  bash setup.sh

set -e
echo ""
echo " GBE NVM Builder — Environment Setup"
echo " ====================================="

if ! command -v python3 &>/dev/null; then
    echo " ERROR: python3 not found. Install Python 3.10+ from https://python.org"
    exit 1
fi

echo " Creating virtual environment (.venv)..."
python3 -m venv .venv

echo " Installing dependencies..."
.venv/bin/python -m pip install --quiet --upgrade pip
.venv/bin/python -m pip install --quiet openpyxl

echo ""
echo " Setup complete!"
echo ""
echo " To build a GBE NVM image, run:"
echo "   ./build.sh"
echo ""
