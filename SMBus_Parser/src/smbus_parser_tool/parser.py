import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class SmbusRecord:
    row_number: int
    timestamp: str
    address: str
    operation: str
    data: str


def parse_csv(path: str | Path) -> list[SmbusRecord]:
    csv_path = Path(path)
    with csv_path.open("r", encoding="utf-8-sig", newline="") as input_file:
        reader = csv.DictReader(input_file)
        if reader.fieldnames is None:
            return []
        return list(records_from_rows(reader))


def records_from_rows(rows: Iterable[dict[str, str]]) -> Iterable[SmbusRecord]:
    for row_number, row in enumerate(rows, start=2):
        yield SmbusRecord(
            row_number=row_number,
            timestamp=first_value(row, "time", "timestamp", "start time"),
            address=first_value(row, "address", "addr", "slave address"),
            operation=first_value(row, "operation", "op", "read/write", "rw"),
            data=first_value(row, "data", "bytes", "raw bytes", "value"),
        )


def first_value(row: dict[str, str], *names: str) -> str:
    normalized = {key.strip().lower(): value.strip() for key, value in row.items() if key is not None and value is not None}
    for name in names:
        value = normalized.get(name)
        if value:
            return value
    return ""


def summarize(records: list[SmbusRecord]) -> str:
    reads = sum(1 for record in records if "r" in record.operation.lower())
    writes = sum(1 for record in records if "w" in record.operation.lower())
    addresses = sorted({record.address for record in records if record.address})
    lines = [
        f"Records: {len(records)}",
        f"Reads: {reads}",
        f"Writes: {writes}",
        f"Unique addresses: {len(addresses)}",
    ]
    if addresses:
        lines.append("Addresses: " + ", ".join(addresses[:20]))
    return "\n".join(lines)
