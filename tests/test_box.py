from pptreport import PowerPointReport
import pandas as pd
import pytest
import shutil
import os

content_dir = "examples/content/"


def test_empty_content():

    report = PowerPointReport()
    report.add_slide([None, "A text"])

    slide = report._slides[0]

    assert len(slide._boxes) == 2
    assert slide._boxes[0].content_type == "empty"
    assert slide._boxes[1].content_type == "text"


def test_estimate_fontsize():
    """ Check that error is correctly raised and caught when fontsize cannot be estimated """

    short_word = "This is a short text".replace(" ", "-")
    long_word = "This is a very long text with a looooooooooooooong word to find fontsize for, but which might give an error".replace(" ", "-")

    report = PowerPointReport()
    report.add_slide([short_word, long_word], width_ratios=[10, 1])  # provoke very little space for long text

    assert len(report._slides) == 1


@pytest.mark.parametrize("fontsize", [12, "12", "big"])
def test_set_fontsize(fontsize):
    """ Test that fontsize is correctly set and validated """

    text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla euismod, nisl sed aliquam lacinia"

    report = PowerPointReport()

    if fontsize == "big":
        with pytest.raises((ValueError), match=None):
            report.add_slide([text], fontsize=fontsize)
    else:
        report.add_slide([text], fontsize=fontsize)
        assert len(report._slides) == 1


@pytest.mark.parametrize("show_borders", [True, False])
def test_borders(show_borders):
    """ Test that borders of boxes are added """

    report = PowerPointReport()
    report.add_slide("A text", show_borders=show_borders)

    if show_borders is True:
        assert report._slides[0]._boxes[0].border is not None
    else:
        assert report._slides[0]._boxes[0].border is None


@pytest.mark.parametrize("content", [content_dir + "wrong_extension.pdf",
                                     content_dir + "wrong_extension.jpg"])
def test_content_fill(content):
    """ Test that contents with false extensions are correctly identified """

    report = PowerPointReport()

    # Create copy of picture to .pdf and vice versa for testing
    if os.path.splitext(content)[-1] == ".pdf":
        shutil.copyfile(content_dir + "cat.jpg", content)
    elif os.path.splitext(content)[-1] == ".jpg":
        shutil.copyfile(content_dir + "chips.pdf", content)

    with pytest.raises(ValueError, match="Could not open"):
        report.add_slide(content=content)

    # Clean up
    os.remove(content)


@pytest.mark.parametrize("content", [pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})])
def test_invalid_content_type(content):
    """ Test input of invalid content types """

    report = PowerPointReport()
    with pytest.raises(ValueError):
        report.add_slide(content=content)


@pytest.mark.parametrize("string, warning",
                         [("**Bold**", None),
                          ("_Italic_", None),
                          ("**Bold and _italic_**", None),  # nested markdown
                          ("# A title", None),
                          ("[Link](a/link)", None),
                          ("`inline code`", None),
                          ("----", "Markdown horizontal rules are not supported"),
                          ("- list entry", "Markdown lists are not supported"),
                          ("   1. list entry", "Markdown lists are not supported"),
                          ("```\ncode block\n```", "Markdown code blocks are not supported"),
                          ("> Blockquote", "Markdown block quotes are not supported"),
                          ("![](image.png)", "Markdown images are not supported")])
def test_markdown_warning(caplog, string, warning):

    report = PowerPointReport()
    report.add_slide(string)

    if warning is not None:
        assert warning in caplog.text
    else:
        assert "not supported" not in caplog.text  # check that no warning about markdown support was raised
