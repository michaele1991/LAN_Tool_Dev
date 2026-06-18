# Driver Collector

Portable helper for collecting Intel LAN driver traces, converting ETL logs to CSV, and keeping the selected driver/flow/script mapping explicit.

## Supported Choices

| Family | Driver | Idle flow | CS flow | Sx flow |
|---|---|---|---|---|
| GBE | NDIS driver | regular `start_trace.bat` / `stop_trace.bat` with driver `E1D` | regular `start_trace.bat` / `stop_trace.bat` with driver `E1D` | not mapped yet: no continuous NDIS script found |
| FXVL | NDIS driver | regular `start_trace.bat` / `stop_trace.bat` with driver `E2F` | regular `start_trace.bat` / `stop_trace.bat` with driver `E2F` | not mapped yet: no continuous NDIS script found |
| GBE | NetAdapter | `StartTrace.bat` / `StopTrace.bat`, verbose config `ConfigWppVerboseE1dn.bat` | `StartTrace.bat` / `StopTrace.bat`, verbose config `ConfigWppVerboseE1dn.bat` | `StartContinuousBootTrace.bat` / `StopBootTrace.bat` |
| FXVL | NetAdapter | `StartTrace.bat` / `StopTrace.bat`, verbose config `ConfigWppVerboseE2fn.bat` | `StartTrace.bat` / `StopTrace.bat`, verbose config `ConfigWppVerboseE2fn.bat` | `StartContinuousBootTrace.bat` / `StopBootTrace.bat` |

## Quick Start

### GUI

```cmd
RUN_APP.bat
```

### CLI

```cmd
RUN_APP.bat plan --family GBE --driver NetAdapter --flow Idle
RUN_APP.bat start --family GBE --driver NetAdapter --flow Idle --tag my_test
RUN_APP.bat stop --family GBE --driver NetAdapter --flow Idle
RUN_APP.bat export-csv --etl C:\Logs\clientLanLog.etl --output C:\Logs\clientLanLog.csv --symbols C:\Symbols\e1dn.pdb
```

## Notes

- Start/stop actions must usually run as Administrator because the underlying scripts use `logman` and registry settings.
- `Load Symbols` records a PDB/symbol path and sets `_NT_SYMBOL_PATH` for child tools. The CSV export uses `tracerpt.exe` when available.
- The app does not silently guess unavailable flows. If no script was found for a selected option, it reports the missing mapping.
