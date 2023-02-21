from pptreport.classes import PowerPointReport
import pytest
import json
import os


@pytest.mark.parametrize("full", [True, False])
def test_config_writing_reading(full):
    """ Test that reading/writing the same config will result in the same report """

    # Create report and save to config
    report1 = PowerPointReport()
    global_params = {"outer_margin": 1, "top_margin": 1.5}
    report1.add_global_parameters(global_params)
    report1.add_slide("A text")                                     # test default
    report1.add_slide(["text1", "text2"], width_ratios=[0.8, 0.2])  # test list
    report1.add_slide(["text1", "text2"], split=True)               # test bool
    report1.write_config("report1.json", full=full)

    # Create new report with config
    report2 = PowerPointReport()
    report2.from_config("report1.json")
    report2.write_config("report2.json", full=full)

    # Assert that the written config is the same
    with open("report1.json", "r") as f:
        config1 = json.load(f)

    with open("report2.json", "r") as f:
        config2 = json.load(f)

    os.remove("report1.json")
    os.remove("report2.json")

    assert config1 == config2


def test_get_config_global():
    """ Test that get_config takes into account the current global parameters """

    # Create report
    report = PowerPointReport()
    global_params = {"outer_margin": 1, "top_margin": 1.5}
    report.add_global_parameters(global_params)

    # Add a slide with same parameters as global
    report.add_slide("A text", **global_params)

    # Change global parameters again
    new_global = {"outer_margin": 0, "top_margin": 2.5}
    report.add_global_parameters(new_global)

    # Add a slide with same parameters as global
    report.add_slide("Another text", **new_global)

    # Create config
    config = report.get_config()

    # Assert that config takes into account that globals were updated
    for key, value in global_params.items():

        # First slide should have old global parameters
        assert config["slides"][0][key] == value

        # Second slide should have no parameter (as these are default values)
        assert key not in config["slides"][1]
