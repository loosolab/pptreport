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


@pytest.mark.parametrize("content", ["examples/content/fish_description.txt",
                                     "examples/content/fish_description.md",
                                     "examples/content/cat.jpg",
                                     "examples/content/chips.pdf"])
def test_content_fill(content):
    """ Test that filling of slides with different types of content does not throw an error """

    report = PowerPointReport(verbosity=2)
    report.add_slide(content=content)

    assert len(report._slides) == 1  # assert that a slide was added


def test_pdf_output(caplog):
    """ Test that pdf output works """

    report = PowerPointReport()
    report.add_slide("A text")
    report.save("test.pptx", pdf=True)

    if caplog.text != "":  # if libreoffice is installed, caplog will be empty
        assert "Option 'pdf' is set to True, but LibreOffice could not be found on path." in caplog.text


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


# even if missing_file == "raise", content that is clearly not a file should not raise an error
def test_missing_file_content():

    content = ["A string which is not a file. Should not raise an error."]
    report = PowerPointReport()
    report.add_slide(content=content, missing_file="raise")


@pytest.mark.parametrize("missing_file", ["raise", "empty", "text", "skip"])
def test_missing_file_option(caplog, missing_file):
    """ Test that a missing file raises an error or just warning """

    report = PowerPointReport(verbosity=1)

    pattern = "examples/content/*.txtt"

    if missing_file == "raise":
        with pytest.raises(FileNotFoundError):
            report.add_slide(pattern, missing_file=missing_file)  # no files found with this extension

    else:

        report.add_slide(pattern, missing_file=missing_file)
        assert "No files could be found for pattern" in caplog.text

        if missing_file == "skip":
            assert len(report._slides[0]._boxes) == 0  # slide added, but no boxes added

        elif missing_file == "empty":
            assert len(report._slides[0]._boxes) == 1  # one box added (empty)
            assert report._slides[0]._boxes[0].content is None

        elif missing_file == "text":
            assert len(report._slides[0]._boxes) == 1  # one box added with string content
            assert report._slides[0]._boxes[0].content == pattern


@pytest.mark.parametrize("empty_slide", ["keep", "skip"])
def test_empty_slide(empty_slide):

    report = PowerPointReport(verbosity=1)

    pattern = "examples/content/*.txtt"  # no files found with this extension
    report.add_slide(pattern, missing_file="text", empty_slide=empty_slide)

    if empty_slide == "keep":
        assert len(report._slides) == 1
    elif empty_slide == "skip":
        assert len(report._slides) == 0  # no slides added because slide was empty
