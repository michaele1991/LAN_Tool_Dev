#!/usr/bin/env bash
# GBE NVM Builder — launch the 5-question build wizard
set -e

VENV=".venv/bin/python"
WIZARD="GBE_Builder/wizard.py"

if [ ! -f "$VENV" ]; then
    echo "Virtual environment not found — running setup first..."
    bash setup.sh
fi

"$VENV" "$WIZARD"
