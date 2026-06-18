#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
export PYTHONPATH="$PWD/src${PYTHONPATH:+:$PYTHONPATH}"

if command -v python3 >/dev/null 2>&1; then
  python3 -m driver_collector_tool "$@"
elif command -v python >/dev/null 2>&1; then
  python -m driver_collector_tool "$@"
else
  echo "Python 3 was not found. Install Python 3.10 or newer." >&2
  exit 1
fi
