from pptreport import PowerPointReport
import pytest

content_dir = "examples/content/"


#####################################################################
# Tests for input to PowerPoint presentation
#####################################################################

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


@pytest.mark.parametrize("verbosity, valid",
                         [(0, True),
                          (1, True),
                          (2, True),
                          (3, False),
                          ("invalid", False),
                          (False, False)])
def test_verbosity(verbosity, valid):
    """ Test that the logger levels are correct """

    if valid:
        report = PowerPointReport(verbosity=verbosity)
    else:
        with pytest.raises(ValueError):
            report = PowerPointReport(verbosity=verbosity)


#####################################################################
# Tests for input to .add_slide
#####################################################################

def test_validation(config, valid, match="Invalid value for "):

    default_config = {"content": ["A text", content_dir + "cat.jpg", content_dir + "chips.pdf"]}
    default_config.update(config)

    report = PowerPointReport()
    if valid:
        report.add_slide(**default_config)
    else:
        with pytest.raises((ValueError, TypeError), match=match) as e:
            report.add_slide(**default_config)

        print(f"Configuration {config} failed with error: {e.value}\n")

# ------------------------------------------------------------------- #
@pytest.mark.parametrize("content, valid",
                         [("A text", True),
                          (content_dir + "cat.jpg", True),
                          ([], True),
                          (1, True)])
def test_content_input(content, valid):
    config = {"content": content}
    test_validation(config, valid)


# ------------------------------------------------------------------- #
# grouped content
@pytest.mark.parametrize("grouped_content, valid",
                         [(["no", "groups"], False),
                          ("A text", False)])
def test_grouped_content(grouped_content, valid):
    config = {"content": None, "grouped_content": grouped_content}
    test_validation(config, valid)


# ------------------------------------------------------------------- #
@pytest.mark.parametrize("title, valid",
                         [("A title", True),
                          (None, True),
                          (1, True),
                          ([], False)])
def test_title_input(title, valid):
    config = {"title": title}
    test_validation(config, valid)

# ------------------------------------------------------------------- #
# slide layout
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


# ------------------------------------------------------------------- #
# content layout
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


# ------------------------------------------------------------------- #
# content alignment
@pytest.mark.parametrize("content_alignment, valid",
                         [("left", True),
                          ("center", True),
                          ("right", True),
                          ("center right", True),
                          ("upper wherever", False),
                          ("invalid", False),
                          (0, False)])
def test_content_alignment(content_alignment, valid):
    config = {"content_alignment": content_alignment}
    test_validation(config, valid)


# ------------------------------------------------------------------- #
# margins
@pytest.mark.parametrize("parameter", ["outer_margin", "inner_margin", "top_margin", "bottom_margin", "left_margin", "right_margin"])
@pytest.mark.parametrize("margins, valid",
                         [(0.1, True),
                          ("1", True),
                          (0, True),
                          (-2, False)])
def test_margins_input(margins, valid, parameter):

    config = {parameter: margins}
    test_validation(config, valid)


# ------------------------------------------------------------------- #
# ratios
@pytest.mark.parametrize("parameter", ["width_ratios", "height_ratios"])
@pytest.mark.parametrize("ratios, valid",
                         [([1, 2], True),
                          ([0, 1], False),
                          ([0, 1, 2, 3], False),
                          ([0, -2], False),
                          (False, False),
                          (0, False)])
def test_ratios_input(ratios, valid, parameter):
    config = {parameter: ratios}
    test_validation(config, valid)


# ------------------------------------------------------------------- #
# notes
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


# ------------------------------------------------------------------- #
# split
@pytest.mark.parametrize("split, valid", 
                         [(True, True),
                          (False, True),
                          (2, True),
                          ("2", True),
                          ("invalid", False),
                          (0, False),
                          ([], False)])
def test_split_input(split, valid):
    config = {"split": split}
    test_validation(config, valid)


# ------------------------------------------------------------------- #
# show filename
@pytest.mark.parametrize("show_filename, valid",
                         [(True, True),
                          ("filename", True),
                          ("filename_ext", True),
                          ("filepath", True),
                          ("filepath_ext", True),
                          ("path", True),
                          ("invalid", False),
                          ([], False)])
def test_show_filename_input(show_filename, valid):
    config = {"show_filename": show_filename}
    test_validation(config, valid)


# ------------------------------------------------------------------- #
#filename alignment
@pytest.mark.parametrize("filename_alignment, valid",
                         [("left", True),
                          ("center", True),
                          ("right", True),
                          ("RIGHT", True),
                          ("right ", True),
                          ("center right", False),
                          ("invalid", False),
                          (0, False)])
def test_filename_alignment(filename_alignment, valid):
    config = {"filename_alignment": filename_alignment, "show_filename": True}
    test_validation(config, valid, "Invalid value for 'filename_alignment'")


# ------------------------------------------------------------------- #
# fill_by
@pytest.mark.parametrize("fill_by, valid",
                         [("row", True),
                          ("column", True),
                          ("invalid", False),
                          (0, False),
                          ([], False)])
def test_fill_by_input(fill_by, valid):
    config = {"fill_by": fill_by}
    test_validation(config, valid)

# ------------------------------------------------------------------- #
# remove_placeholders
@pytest.mark.parametrize("remove_placeholders, valid",
                         [(True, True),
                          (False, True),
                          ("True", True),
                          ("invalid", False),
                          ([], False)])
def test_remove_placeholders_input(remove_placeholders, valid):
    config = {"remove_placeholders": remove_placeholders}
    test_validation(config, valid, "Invalid input for ")


# ------------------------------------------------------------------- #
# fontsize
@pytest.mark.parametrize("fontsize, valid",
                         [(1, True),
                          ("1", True),
                          (0, False),
                          (-2, False),
                          ("big", False),
                          ([], False)])
def test_fontsize_input(fontsize, valid):
    config = {"fontsize": fontsize}
    test_validation(config, valid, "Invalid input for ")

# ------------------------------------------------------------------- #
# pdf_pages
@pytest.mark.parametrize("pdf_pages, valid",
                         [(1, True),
                          ("1", True),
                          ("all", True),
                          ("1,2", True),
                          ([1, 2, 2], True),
                          ("invalid", False),
                          (0, False),
                          ([0], False)])
def test_pdfpages_input(pdf_pages, valid):
    config = {"pdf_pages": pdf_pages, "content": content_dir + "pdfs/multidogs_1.pdf"}
    test_validation(config, valid, "Invalid input for ")


@pytest.mark.parametrize("pdf_pages, valid", [("all", False),
                                              ([1, 2], False),
                                              (1, True)])
def test_pdf_pages_grouped(pdf_pages, valid):

    config = {"content": None, "grouped_content": [content_dir + "pdfs/multidogs_([0-9]).pdf"], "pdf_pages": pdf_pages}
    test_validation(config, valid, "Invalid input for ")


# ------------------------------------------------------------------- #
# missing_file
@pytest.mark.parametrize("missing_file, valid",
                         [("raise", True),
                          ("empty", True),
                          ("text", True),
                          ("skip", True),
                          ("invalid", False),
                          (True, False)])
def test_missing_file_input(missing_file, valid):
    config = {"missing_file": missing_file}
    test_validation(config, valid, "Invalid input for ")

# ------------------------------------------------------------------- #
# empty_slide
@pytest.mark.parametrize("empty_slide, valid",
                         [("keep", True),
                          ("skip", True),
                          ("invalid", False),
                          (False, False)])
def test_empty_slide(empty_slide, valid):
    config = {"empty_slide": empty_slide}
    test_validation(config, valid, "Invalid input for ")


# ------------------------------------------------------------------- #
# integers
@pytest.mark.parametrize("parameter", ["n_columns", "dpi", "max_pixels"])
@pytest.mark.parametrize("value, valid",
                         [(1, True),
                          ("1e3", True),
                          ("2", True),
                          (0, False),
                          (-1, False),
                          ("invalid", False),
                          (False, False)])
def test_integers(value, valid, parameter):
    config = {parameter: value}
    test_validation(config, valid, match=f"Invalid value for '{parameter}' parameter")


# ------------------------------------------------------------------- #
# max_pixels
@pytest.mark.parametrize("value, valid",
                         [("1e3", True),
                          (4, True),
                          (1, False),
                          ("2", False),
                          (0, False),
                          (-1, False),
                          ("invalid", False),
                          (False, False)])
def test_max_pixels_input(value, valid):
    config = {"max_pixels": value}
    test_validation(config, valid, match=f"Invalid value for 'max_pixels' parameter")


# ------------------------------------------------------------------- #
# invalid combinations
@pytest.mark.parametrize("params", [{"content": ["text"], "grouped_content": ["text"]},    # both content and grouped_content given
                                    {"content": None, "split": True}])   # content has to given when split is True
def test_combinations(params):

    test_validation(params, False, "Invalid input combination")
