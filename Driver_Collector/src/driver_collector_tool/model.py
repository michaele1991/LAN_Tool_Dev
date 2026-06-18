from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class CollectorSelection:
    family: str
    driver: str
    flow: str


@dataclass(frozen=True)
class ScriptPlan:
    selection: CollectorSelection
    supported: bool
    start_script: Path | None
    stop_script: Path | None
    parser_hint: str
    driver_code: str | None = None
    config_script: Path | None = None
    notes: str = ""

    def describe(self) -> str:
        lines = [
            f"Family: {self.selection.family}",
            f"Driver: {self.selection.driver}",
            f"Flow: {self.selection.flow}",
            f"Supported: {'yes' if self.supported else 'no'}",
        ]
        if self.driver_code:
            lines.append(f"Driver code: {self.driver_code}")
        if self.config_script:
            lines.append(f"Verbose config script: {self.config_script}")
        if self.start_script:
            lines.append(f"Start script: {self.start_script}")
        if self.stop_script:
            lines.append(f"Stop script: {self.stop_script}")
        lines.append(f"Parse/export: {self.parser_hint}")
        if self.notes:
            lines.append(f"Notes: {self.notes}")
        return "\n".join(lines)
