from unittest.mock import patch
from driver_collector_tool.config import build_plan
from driver_collector_tool.runner import start_collection, stop_collection


# ── config / plan tests ────────────────────────────────────────────────────

def test_netadapter_idle_uses_regular_scripts():
    plan = build_plan("GBE", "NetAdapter", "Idle flow")
    assert plan.supported
    assert plan.start_script.name == "StartTrace.bat"
    assert plan.stop_script.name == "StopTrace.bat"
    assert plan.config_script.name == "ConfigWppVerboseE1dn.bat"


def test_netadapter_sx_uses_continuous_boot_scripts():
    plan = build_plan("FXVL", "NetAdapter", "Sx flow")
    assert plan.supported
    assert plan.start_script.name == "StartContinuousBootTrace.bat"
    assert plan.stop_script.name == "StopBootTrace.bat"
    assert plan.config_script.name == "ConfigWppVerboseE2fn.bat"


def test_ndis_sx_uses_boot_scripts():
    plan = build_plan("GBE", "NDIS driver", "Sx flow")
    assert plan.supported
    assert plan.start_script.name == "start_boot_trace.bat"
    assert plan.stop_script.name  == "stop_boot_trace.bat"
    assert plan.driver_code == "E1D"


def test_ndis_fxvl_sx_uses_boot_scripts():
    plan = build_plan("FXVL", "NDIS driver", "Sx flow")
    assert plan.supported
    assert plan.start_script.name == "start_boot_trace.bat"
    assert plan.stop_script.name  == "stop_boot_trace.bat"
    assert plan.driver_code == "E2F"


# ── runner argument-building tests ────────────────────────────────────────
# Patch run_script so no actual bat files are executed.

def _capture_args(captured: list):
    """Return a run_script replacement that records (script, args)."""
    def _fake(script, args):
        captured.append((script, args))
        return "ok"
    return _fake


def test_ndis_idle_start_no_tag_passes_driver_only():
    plan = build_plan("GBE", "NDIS driver", "Idle flow")
    captured = []
    with patch("driver_collector_tool.runner.run_script", _capture_args(captured)):
        start_collection(plan, tag="")
    _, args = captured[0]
    assert args == ["E1D"]


def test_ndis_idle_start_with_tag_uses_explicit_defaults():
    """Empty-string args must NOT be passed — they cause cmd.exe quoting errors."""
    plan = build_plan("GBE", "NDIS driver", "Idle flow")
    captured = []
    with patch("driver_collector_tool.runner.run_script", _capture_args(captured)):
        start_collection(plan, tag="test_001")
    _, args = captured[0]
    assert args == ["E1D", "0x0000003F", "4", "test_001"]
    assert "" not in args  # no empty strings that become "" in cmd.exe


def test_ndis_idle_stop_with_tag_passes_full_session_name():
    """stop_trace.bat expects the full session name, not driver + tag separately."""
    plan = build_plan("GBE", "NDIS driver", "Idle flow")
    captured = []
    with patch("driver_collector_tool.runner.run_script", _capture_args(captured)):
        stop_collection(plan, tag="test_001")
    _, args = captured[0]
    assert args == ["E1D_test_001"]


def test_ndis_idle_stop_no_tag_passes_driver_code():
    plan = build_plan("GBE", "NDIS driver", "Idle flow")
    captured = []
    with patch("driver_collector_tool.runner.run_script", _capture_args(captured)):
        stop_collection(plan, tag="")
    _, args = captured[0]
    assert args == ["E1D"]


def test_ndis_sx_stop_with_tag_passes_driver_and_tag_separately():
    """stop_boot_trace.bat reconstructs NdisBootLog_<driver>_<tag> from two args."""
    plan = build_plan("GBE", "NDIS driver", "Sx flow")
    captured = []
    with patch("driver_collector_tool.runner.run_script", _capture_args(captured)):
        stop_collection(plan, tag="sx_repro")
    _, args = captured[0]
    assert args == ["E1D", "sx_repro"]


def test_netadapter_start_with_tag_passes_tag():
    plan = build_plan("GBE", "NetAdapter", "Idle flow")
    captured = []
    with patch("driver_collector_tool.runner.run_script", _capture_args(captured)):
        start_collection(plan, tag="my_run")
    _, args = captured[0]
    assert args == ["my_run"]

