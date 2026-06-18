from pathlib import Path

from .model import CollectorSelection, ScriptPlan


TOOL_ROOT = Path(__file__).resolve().parents[2]
LAN_REPO_ROOT = TOOL_ROOT.parent
REGRESSION_ROOT = LAN_REPO_ROOT.parent

FAMILIES = ("GBE", "FXVL")
DRIVERS = ("NDIS driver", "NetAdapter")
FLOWS = ("Idle flow", "CS flow", "Sx flow")

NETADAPTER_SCRIPT_DIR = REGRESSION_ROOT / "07_TOOLS" / "Driver_NDIS_NetAdapter_collectors" / "NetAdapter" / "ClienLanTrace" / "ClienLanTrace"
NDIS_SCRIPT_DIR = REGRESSION_ROOT / "drivers.ethernet.windows.ndis-ccv1-master"

FAMILY_DRIVER_CODE = {
    ("GBE", "NDIS driver"): "E1D",
    ("FXVL", "NDIS driver"): "E2F",
}

NETADAPTER_CONFIG_SCRIPT = {
    "GBE": NETADAPTER_SCRIPT_DIR / "ConfigWppVerboseE1dn.bat",
    "FXVL": NETADAPTER_SCRIPT_DIR / "ConfigWppVerboseE2fn.bat",
}


def build_plan(family: str, driver: str, flow: str) -> ScriptPlan:
    selection = CollectorSelection(family=family, driver=driver, flow=flow)
    if family not in FAMILIES:
        return unsupported(selection, f"Unknown family: {family}")
    if driver not in DRIVERS:
        return unsupported(selection, f"Unknown driver: {driver}")
    if flow not in FLOWS:
        return unsupported(selection, f"Unknown flow: {flow}")

    if driver == "NetAdapter":
        if flow in ("Idle flow", "CS flow"):
            return ScriptPlan(
                selection=selection,
                supported=True,
                start_script=NETADAPTER_SCRIPT_DIR / "StartTrace.bat",
                stop_script=NETADAPTER_SCRIPT_DIR / "StopTrace.bat",
                config_script=NETADAPTER_CONFIG_SCRIPT[family],
                parser_hint="Use tracerpt.exe to export clientLanLog.etl to CSV after StopTrace completes.",
                notes="Regular/common trace flow. The scripts create CLIENT_LAN_LOG_* folders and collect environment snapshots.",
            )
        return ScriptPlan(
            selection=selection,
            supported=True,
            start_script=NETADAPTER_SCRIPT_DIR / "StartContinuousBootTrace.bat",
            stop_script=NETADAPTER_SCRIPT_DIR / "StopBootTrace.bat",
            config_script=NETADAPTER_CONFIG_SCRIPT[family],
            parser_hint="Use tracerpt.exe to export ClientLanLog.etl from the CLIENT_LAN_CONT_BOOT_LOG_* folder to CSV after StopBootTrace.",
            notes="Sx/reboot flow uses continuous boot logging because regular trace sessions stop across reboot/Sx transitions.",
        )

    driver_code = FAMILY_DRIVER_CODE.get((family, driver))
    if flow in ("Idle flow", "CS flow"):
        return ScriptPlan(
            selection=selection,
            supported=True,
            start_script=NDIS_SCRIPT_DIR / "start_trace.bat",
            stop_script=NDIS_SCRIPT_DIR / "stop_trace.bat",
            driver_code=driver_code,
            parser_hint="Use tracerpt.exe to export the generated <driver>_<tag>.etl file to CSV.",
            notes="Regular WPP logman flow. Start command passes the driver code to start_trace.bat.",
        )

    return unsupported(
        selection,
        "No NDIS continuous/Sx collector script was found. Existing NDIS package contains only start_trace.bat and stop_trace.bat; add boot/continuous scripts to enable Sx for NDIS.",
        driver_code=driver_code,
    )


def unsupported(selection: CollectorSelection, reason: str, driver_code: str | None = None) -> ScriptPlan:
    return ScriptPlan(
        selection=selection,
        supported=False,
        start_script=None,
        stop_script=None,
        driver_code=driver_code,
        parser_hint="Not available until the missing collector script is provided.",
        notes=reason,
    )
