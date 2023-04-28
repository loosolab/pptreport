from pptreport import PowerPointReport
import pytest

content_dir = "examples/content/"


def test_empty_content():

    report = PowerPointReport(verbosity=2)
    report.add_slide([None, "A text"])

    slide = report._slides[0]

    assert len(slide._boxes) == 2
    assert slide._boxes[0].content_type == "empty"
    assert slide._boxes[1].content_type == "text"


@pytest.mark.parametrize("content_alignment", ["left", "center", "right"])
def test_content_alignment(content_alignment):
    """ Test that content alignment is correctly validated """

    pass


def test_estimate_fontsize():
    """ Check that error is correctly raised and caught when fontsize cannot be estimated """

    short_word = "This is a short text".replace(" ", "-")
    long_word = "This is a very long text to find fontsize for, but which might give an error".replace(" ", "-")

    report = PowerPointReport()
    report.add_slide([short_word, long_word])

    assert len(report._slides) == 1


@pytest.mark.parametrize("fontsize", [12, "12", "big"])
def test_set_fontsize(fontsize):
    """ Test that fontsize is correctly set and validated """

    text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla euismod, nisl sed aliquam lacinia"

    report = PowerPointReport()

    if fontsize == "big":
        with pytest.raises(ValueError):
            report.add_slide([text], fontsize=fontsize)
    else:
        report.add_slide([text], fontsize=fontsize)
        assert len(report._slides) == 1
