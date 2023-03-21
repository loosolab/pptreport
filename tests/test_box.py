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
