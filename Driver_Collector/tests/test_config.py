from driver_collector_tool.config import build_plan


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
