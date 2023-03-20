from pptreport import PowerPointReport
import pytest

content_dir = "examples/content/"


def test_title_slide():
    """ Test that title slide is added correctly """

    report = PowerPointReport()
    report.add_title_slide(title="Title", subtitle="Subtitle")
    slide = report._slides[0]

    assert slide._slide.shapes[0].text == "Title"
    assert slide._slide.shapes[1].text == "Subtitle"


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


@pytest.mark.parametrize("options", [{"content": ["text"], "grouped_content": ["text"]},    # both content and grouped_content given
                                     {"content": None, "split": True},   # content has to given when split is True
                                     {"n_columns": "a lot"},
                                     {"show_filename": "invalid"},
                                     {"split": "invalid"}])
def test_invalid_input(options):
    """ Test that invalid input raises ValueError"""

    report = PowerPointReport()
    with pytest.raises(ValueError):

        if "content" not in options and "grouped_content" not in options:
            options["content"] = content_dir + "cat.jpg"

        report.add_slide(**options)


@pytest.mark.parametrize("show_filename", [True, False, "True", "False"])
def test_show_filename(show_filename):
    """ Assert that filenames are added (or not) to the slide """

    report = PowerPointReport()
    report.add_slide(content_dir + "cat.jpg", show_filename=show_filename)  # remove_placeholders=True)

    slide = report._slides[0]
    n_placeholders = len(slide._slide.placeholders)
    if slide.show_filename:
        assert len(slide._slide.shapes) - n_placeholders == 2
        assert slide._slide.shapes[-1].text == content_dir + "cat.jpg"
    else:
        assert len(slide._slide.shapes) - n_placeholders == 1
