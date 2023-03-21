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


@pytest.mark.parametrize("size, valid", [("standard", True),
                                         ("widescreen", True),
                                         ("a4-portrait", True),
                                         ("a4-landscape", True),
                                         ((10, 10), True),
                                         (("10", "10"), True),
                                         ((10, 10, 10), False),
                                         ("invalid", False)])
def test_set_size(size, valid):
    """ Test that set_size works """

    report = PowerPointReport()

    if valid:
        report.set_size(size)
    else:
        with pytest.raises(ValueError):
            report.set_size(size)


def test_borders():
    """ Test that borders of boxes can be added and removed from all slides"""

    report = PowerPointReport()
    report.add_slide("A text")

    # Add borders
    report.add_borders()
    assert report._slides[0]._boxes[0].border is not None

    # Remove borders
    report.remove_borders()
    assert report._slides[0]._boxes[0].border is None


@pytest.mark.parametrize("verbosity", [0, 1, 2])
def test_logger(capfd, verbosity):
    """ Test that the logger levels are correct """

    report = PowerPointReport(verbosity=verbosity)
    report.add_slide("A text")
    out, _ = capfd.readouterr()

    if verbosity == 0:
        assert out == ""
    elif verbosity == 1:
        assert "[INFO]" in out and "[DEBUG]" not in out
    elif verbosity == 2:
        assert "[INFO]" in out and "[DEBUG]" in out


def test_logger_invalid():
    """ Test that an invalid verbosity level raises an error """

    with pytest.raises(ValueError):
        _ = PowerPointReport(verbosity=3)


@pytest.mark.parametrize("slide_layout, valid", [("Title Slide", True),
                                                 (0, True),
                                                 ("Invalid slide", False),  # Invalid slide name
                                                 (100, False),  # Invalid slide number
                                                 ([""], False)  # Invalid type
                                                 ])
def test_slide_layout(slide_layout, valid):
    """ Test that slide_layout is correctly validated """
    report = PowerPointReport()

    if valid:
        report.add_slide("A text", slide_layout=slide_layout)

    else:
        with pytest.raises(Exception):
            report.add_slide("A text", slide_layout=slide_layout)


@pytest.mark.parametrize("notes, valid", [("A note", True),
                                          (["A note", "Another note"], True),
                                          ("examples/content/fish_description.txt", True),
                                          (dict, False),
                                          ([dict], False)
                                          ])
def test_add_notes(notes, valid):
    """ Test that notes can be added to slides, and that an error is thrown if the notes are invalid """

    report = PowerPointReport()

    if valid:
        report.add_slide("A text", notes=notes)
    else:
        with pytest.raises(ValueError, match="Notes must be either a string or a list of strings."):
            report.add_slide("A text", notes=notes)


@pytest.mark.parametrize("content, valid", [("grid", True),
                                            ("vertical", True),
                                            ("horizontal", True),
                                            ([0, 1, 2], True),
                                            ([[0, 1], [2, 3]], True),
                                            ("invalid", False),           # invalid string
                                            ([[0, 1, 2], [3, 4]], False)  # inconsistent number of columns
                                            ])
def test_content_layout(content, valid):
    """ Test that content layout is correctly validated """

    report = PowerPointReport()

    if valid:
        report.add_slide("A text", content_layout=content)
    else:
        if isinstance(content, str):
            with pytest.raises(ValueError, match="Unknown layout string:"):
                report.add_slide("A text", content_layout=content)
        else:
            with pytest.raises(ValueError):
                report.add_slide("A text", content_layout=content)


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
