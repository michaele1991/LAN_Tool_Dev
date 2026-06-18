import csv
import glob
import os
import subprocess
from pathlib import Path


class ExportError(RuntimeError):
    pass


# WDK tracefmt — preferred when a PDB is supplied (gives decoded WPP messages).
_TRACEFMT_SEARCH = [
    r"C:\Program Files (x86)\Windows Kits\10\bin\*\x64\tracefmt.exe",
    r"C:\Program Files\Windows Kits\10\bin\*\x64\tracefmt.exe",
]


def _find_tracefmt() -> Path | None:
    for pattern in _TRACEFMT_SEARCH:
        matches = sorted(glob.glob(pattern), reverse=True)   # newest SDK first
        if matches:
            return Path(matches[0])
    return None


def export_etl_to_csv(
    etl_path: str | Path,
    output_csv: str | Path,
    symbols: str | Path | None = None,
) -> Path:
    etl = Path(etl_path)
    output = Path(output_csv)
    if not etl.exists():
        raise ExportError(f"ETL file was not found: {etl}")
    output.parent.mkdir(parents=True, exist_ok=True)

    tracefmt = _find_tracefmt()
    if symbols and tracefmt and tracefmt.exists():
        _export_via_tracefmt(etl, output, Path(symbols), tracefmt)
    else:
        _export_via_tracerpt(etl, output, str(symbols) if symbols else None)

    return output


# ── tracefmt (WDK, WPP-decoded) ───────────────────────────────────────────────

def _export_via_tracefmt(etl: Path, output_csv: Path, pdb: Path, tracefmt: Path) -> None:
    """Decode ETL with WPP symbols via tracefmt, then normalise to clean CSV."""
    tmp_txt = output_csv.with_suffix(".tracefmt.tmp")
    env = os.environ.copy()

    cmd = [
        str(tracefmt), str(etl),
        "-pdb", str(pdb),
        "-o", str(tmp_txt),
        "-csv", "-csvheader", "-sortableTime", "-nosummary",
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=300)
    except FileNotFoundError as exc:
        raise ExportError(f"tracefmt.exe not found at {tracefmt}") from exc
    except subprocess.TimeoutExpired as exc:
        raise ExportError("tracefmt.exe timed out.") from exc

    if not tmp_txt.exists():
        tmp_txt.unlink(missing_ok=True)
        _export_via_tracerpt(etl, output_csv, str(pdb))
        return

    _convert_tracefmt_csv(tmp_txt, output_csv)
    tmp_txt.unlink(missing_ok=True)


def _convert_tracefmt_csv(src: Path, dst: Path) -> None:
    """
    Normalise the raw tracefmt CSV.
    tracefmt -csv -csvheader columns: EventName, Type, TimeStamp, ThreadID, ProcessorNumber, UserData
    Output columns:                   Provider,  Level, Time,      TID,      CPU,             Message
    """
    RENAME = {
        "TimeStamp": "Time", "ProcessorNumber": "CPU", "ThreadID": "TID",
        "Type": "Level", "UserData": "Message", "EventName": "Provider",
    }
    PREFERRED = ["Time", "CPU", "TID", "Level", "Provider", "Message"]

    try:
        with src.open("r", encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.DictReader(fh))
    except UnicodeDecodeError:
        with src.open("r", encoding="latin-1", newline="") as fh:
            rows = list(csv.DictReader(fh))

    if not rows:
        with dst.open("w", encoding="utf-8", newline="") as fh:
            csv.writer(fh).writerow(PREFERRED)
        return

    orig_keys = list(rows[0].keys())
    col_map   = {k: RENAME.get(k, k) for k in orig_keys}
    all_cols  = [col_map[k] for k in orig_keys]
    ordered   = [c for c in PREFERRED if c in all_cols] + [c for c in all_cols if c not in PREFERRED]

    with dst.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=ordered)
        writer.writeheader()
        for row in rows:
            mapped = {col_map[k]: v for k, v in row.items()}
            writer.writerow({k: mapped.get(k, "") for k in ordered})


# ── tracerpt (built-in Windows, no WPP decode) ────────────────────────────────

def _export_via_tracerpt(etl: Path, output_csv: Path, symbols: str | None) -> None:
    env = os.environ.copy()
    if symbols:
        existing = env.get("_NT_SYMBOL_PATH", "")
        env["_NT_SYMBOL_PATH"] = symbols if not existing else f"{symbols};{existing}"

    cmd = ["tracerpt.exe", str(etl), "-o", str(output_csv), "-of", "CSV", "-y"]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=300)
    except FileNotFoundError as exc:
        raise ExportError("tracerpt.exe was not found. Run on Windows with tracerpt available.") from exc
    except subprocess.TimeoutExpired as exc:
        raise ExportError("tracerpt.exe timed out.") from exc

    if result.returncode != 0:
        raise ExportError((result.stderr or result.stdout or "tracerpt.exe failed").strip())
    if not output_csv.exists():
        raise ExportError(f"tracerpt.exe completed but did not create: {output_csv}")
    normalize_csv(output_csv)


def normalize_csv(path: Path) -> None:
    """Strip BOM and re-write as clean UTF-8 CSV."""
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.reader(fh))
    except UnicodeDecodeError:
        return
    if not rows:
        return
    with path.open("w", encoding="utf-8", newline="") as fh:
        csv.writer(fh).writerows(rows)
