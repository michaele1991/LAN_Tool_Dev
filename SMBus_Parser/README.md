# SMBus Parser Tool

Portable SMBus parser project scaffold for Windows and Linux.

## Goals

- Parse SMBus analyzer CSV exports.
- Provide a user-friendly Tkinter GUI.
- Provide a CLI for batch parsing and automation.
- Keep dependencies minimal for portability.

## Quick Start

### Windows

```cmd
RUN_APP.bat
RUN_APP.bat --input trace.csv --summary
```

### Linux

```bash
./run_app.sh
./run_app.sh --input trace.csv --summary
```

## Project Layout

```text
src/smbus_parser_tool/
  app.py       GUI entry point
  cli.py       command line interface
  parser.py    CSV parsing helpers
```

## Notes

This is a clean project starting point. Existing parser logic can be migrated into `src/smbus_parser_tool/parser.py` and GUI flows into `app.py`.
