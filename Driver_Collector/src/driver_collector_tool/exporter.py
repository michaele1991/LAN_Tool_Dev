import csv
import os
import subprocess
from pathlib import Path


class ExportError(RuntimeError):
    pass


def export_etl_to_csv(etl_path: str | Path, output_csv: str | Path, symbols: str | Path | None = None) -> Path:
    etl = Path(etl_path)
    output = Path(output_csv)
    if not etl.exists():
        raise ExportError(f"ETL file was not found: {etl}")
    output.parent.mkdir(parents=True, exist_ok=True)

    env = os.environ.copy()
    if symbols:
        symbol_path = str(Path(symbols))
        existing = env.get("_NT_SYMBOL_PATH", "")
        env["_NT_SYMBOL_PATH"] = symbol_path if not existing else f"{symbol_path};{existing}"

    command = ["tracerpt.exe", str(etl), "-o", str(output), "-of", "CSV", "-y"]
    try:
        result = subprocess.run(command, capture_output=True, text=True, env=env, timeout=300)
    except FileNotFoundError as exc:
        raise ExportError("tracerpt.exe was not found. Install/enable Windows trace tools or run on Windows with tracerpt available.") from exc
    except subprocess.TimeoutExpired as exc:
        raise ExportError("tracerpt.exe timed out while exporting the ETL file.") from exc

    if result.returncode != 0:
        raise ExportError((result.stderr or result.stdout or "tracerpt.exe failed").strip())
    if not output.exists():
        raise ExportError(f"tracerpt.exe completed but did not create: {output}")
    normalize_csv(output)
    return output


def normalize_csv(path: Path) -> None:
    # Keep the file as a proper CSV even when tracerpt emits odd BOM/header spacing.
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as input_file:
            rows = list(csv.reader(input_file))
    except UnicodeDecodeError:
        return
    if not rows:
        return
    with path.open("w", encoding="utf-8", newline="") as output_file:
        writer = csv.writer(output_file)
        writer.writerows(rows)
