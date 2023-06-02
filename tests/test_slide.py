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


@pytest.mark.parametrize("fill_by", ["row", "column"])
def test_fill_by(fill_by):
    """ Test that content is filled correctly by row or column """

    report = PowerPointReport()
    content = ["text" + str(i + 1) for i in range(4)]

    report.add_slide(content, fill_by=fill_by)
    slide = report._slides[0]

    # Assert using the locations of the boxes
    if fill_by == "row":
        assert slide._boxes[0].top == slide._boxes[1].top
    elif fill_by == "column":
        assert slide._boxes[0].left == slide._boxes[1].left


@pytest.mark.parametrize("show_filename, expected", [(True, "cat"),
                                                     ("True", "cat"),
                                                     ("filename", "cat"),
                                                     ("filename_ext", "cat.jpg"),
                                                     ("filepath", content_dir + "cat"),
                                                     ("filepath_ext", content_dir + "cat.jpg"),
                                                     ("path", content_dir[:-1]),  # remove last slash
                                                     (False, None),
                                                     ("False", None)])
def test_show_filename(show_filename, expected):
    """ Assert that filenames are added (or not) to the slide """

    report = PowerPointReport()
    report.add_slide(content_dir + "cat.jpg", show_filename=show_filename, remove_placeholders=True)
    slide = report._slides[0]

    if isinstance(expected, str):
        assert len(slide._slide.shapes) == 2
        assert slide._slide.shapes[-1].text == expected

    else:
        assert len(slide._slide.shapes) == 1


@pytest.mark.parametrize("config", [{"left_margin": 50}, {"top_margin": 50}])
def test_large_margins(config):
    """ Check that large margins throw an error correctls """

    report = PowerPointReport()
    with pytest.raises(ValueError, match="The .+ of content is negative."):
        report.add_slide(content_dir + "cat.jpg", **config)
