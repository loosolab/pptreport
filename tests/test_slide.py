from pptreport import PowerPointReport
import pytest


@pytest.mark.parametrize("fill_by, valid", [("row", True),
                                            ("column", True),
                                            ("invalid", False)])
def test_fill_by(fill_by, valid):
    """ Test that content is filled correctly by row or column """

    report = PowerPointReport()
    content = ["text" + str(i + 1) for i in range(4)]

    if valid:
        report.add_slide(content, fill_by=fill_by)
        slide = report._slides[0]

        # Assert using the locations of the boxes
        if fill_by == "row":
            assert slide._boxes[0].top == slide._boxes[1].top
        elif fill_by == "column":
            assert slide._boxes[0].left == slide._boxes[1].left

    else:
        with pytest.raises(ValueError):
            report.add_slide(content, fill_by=fill_by)
