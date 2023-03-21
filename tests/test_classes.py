from pptreport import PowerPointReport
import pytest
import json
import os

content_dir = "examples/content/"


@pytest.mark.parametrize("full", [True, False])
@pytest.mark.parametrize("expand", [True, False])
def test_config_writing_reading(full, expand):
    """ Test that reading/writing the same config will result in the same report """

    # Create report and save to config
    report1 = PowerPointReport()
    global_params = {"outer_margin": 1, "top_margin": 1.5}
    report1.add_global_parameters(global_params)
    report1.add_slide("A text")                                     # test default
    report1.add_slide(["text1", "text2"], width_ratios=[0.8, 0.2])  # test list
    report1.add_slide(["text1", "text2"], split=True)               # test bool
    report1.write_config("report1.json", full=full, expand=expand)

    # Create new report with config
    report2 = PowerPointReport()
    report2.from_config("report1.json")
    report2.write_config("report2.json", full=full, expand=expand)

    # Assert that the written config is the same
    with open("report1.json", "r") as f:
        config1 = json.load(f)

    with open("report2.json", "r") as f:
        config2 = json.load(f)

    os.remove("report1.json")
    os.remove("report2.json")

    assert config1 == config2


def test_get_config_global():
    """ Test that get_config takes into account the current global parameters, in case they were added multiple times """

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


@pytest.mark.parametrize("expand", [True, False])
def test_get_config(expand):
    """ Test that get_config returns the correct config """

    report = PowerPointReport()
    report.add_slide(content_dir + "*_fish.jpg")

    config = report.get_config(expand=expand)

    if expand is True:
        assert isinstance(config["slides"][0]["content"], list)
        assert len(config["slides"][0]["content"]) == 3
    else:
        assert isinstance(config["slides"][0]["content"], str)  # not expanded


@pytest.mark.parametrize("pdf_pages", ["all", 2, [1, 3], [2, 3, 3]])
@pytest.mark.parametrize("max_allowed", [None, 3])
def test_max_pdf_pages_pass(pdf_pages, max_allowed):

    report = PowerPointReport()
    tmp_files = report.convert_pdf(content_dir + "pdfs/multidogs.pdf", pdf_pages, max_allowed)
    for tmp_file in tmp_files:
        os.remove(tmp_file)


@pytest.mark.parametrize("pdf_pages", ["all", [1, 3], [2, 3, 3]])
@pytest.mark.parametrize("max_allowed", [1])
def test_max_pdf_pages_error(pdf_pages, max_allowed):
    with pytest.raises(ValueError):
        report = PowerPointReport()
        report.convert_pdf(content_dir + "pdfs/multidogs.pdf", pdf_pages, max_allowed)


@pytest.mark.parametrize("pdf_pages", [None, "h", -1, 0, 4])
def test_index_pdf_pages_error(pdf_pages):
    with pytest.raises(IndexError):
        report = PowerPointReport()
        report.convert_pdf(content_dir + "pdfs/multidogs.pdf", pdf_pages)
