import os
import subprocess
from pathlib import Path

from .model import ScriptPlan


class CollectorError(RuntimeError):
    pass


def run_config(plan: ScriptPlan) -> str:
    if not plan.config_script:
        return "No verbose config script is mapped for this selection."
    return run_script(plan.config_script, [])


def start_collection(plan: ScriptPlan, tag: str = "") -> str:
    ensure_supported(plan)
    if plan.start_script is None:
        raise CollectorError("No start script is mapped.")
    args: list[str] = []
    if plan.selection.driver == "NDIS driver":
        if not plan.driver_code:
            raise CollectorError("No NDIS driver code is mapped.")
        args = [plan.driver_code]
        if tag:
            # start_trace.bat signature: <driver> [flags] [level] [session_name]
            # Use explicit defaults instead of "" to avoid cmd.exe quoting issues
            # where set flags=%2 would assign literal "" (two quote chars) not empty.
            args += ["0x0000003F", "4", tag]
    elif tag:
        args = [tag]
    return run_script(plan.start_script, args)


def stop_collection(plan: ScriptPlan, tag: str = "") -> str:
    ensure_supported(plan)
    if plan.stop_script is None:
        raise CollectorError("No stop script is mapped.")
    args: list[str] = []
    if plan.selection.driver == "NDIS driver" and plan.driver_code:
        if plan.selection.flow == "Sx flow":
            # stop_boot_trace.bat accepts <driver> [tag] and reconstructs
            # the autosession name NdisBootLog_<driver>[_tag].
            args = [plan.driver_code]
            if tag:
                args.append(tag)
        else:
            # stop_trace.bat expects the full session name (e.g. E1D or E1D_test_001).
            # start_trace.bat builds: session = driver + "_" + session_name
            session = plan.driver_code + (f"_{tag}" if tag else "")
            args = [session]
    return run_script(plan.stop_script, args)


def ensure_supported(plan: ScriptPlan) -> None:
    if not plan.supported:
        raise CollectorError(plan.notes or "Selection is not supported.")


def run_script(script: Path, args: list[str]) -> str:
    if not script.exists():
        raise CollectorError(f"Script was not found: {script}")
    command = subprocess.list2cmdline([str(script), *args])
    env = os.environ.copy()
    try:
        result = subprocess.run(command, cwd=str(script.parent), capture_output=True, text=True, shell=True, env=env, timeout=30)
    except subprocess.TimeoutExpired:
        return f"Started or still running: {script.name}\nThe script did not finish within the short UI timeout. Check the script/admin prompt window."
    output = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
    if result.returncode != 0:
        raise CollectorError(output or f"Script failed with exit code {result.returncode}: {script}")
    return output or f"Completed: {script.name}"
