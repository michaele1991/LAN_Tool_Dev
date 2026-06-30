#!/usr/bin/env bash
# GBE NVM Builder — launch the 5-question build wizard
set -e

cd "$(dirname "$0")"
VENV=".venv/bin/python"
WIZARD="src/wizard.py"

if [ ! -f "$VENV" ]; then
    echo "Virtual environment not found — running setup first..."
    bash setup.sh
fi

"$VENV" "$WIZARD"
