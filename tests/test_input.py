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
        _ = PowerPointReport(verbosity=verbosity)
    else:
        with pytest.raises(ValueError):
            _ = PowerPointReport(verbosity=verbosity)


@pytest.mark.parametrize("global_parameters, valid",
                         [({"outer_margin": 2}, True),
                          ({"not_valid": 2}, False),
                          (["alist"], False)  # must be dict
                          ])
def test_global_parameters(global_parameters, valid):
    """ Test that global parameters are validated correctly """
    report = PowerPointReport()
    if valid:
        report.add_global_parameters(global_parameters)
    else:
        with pytest.raises((ValueError, TypeError)):
            report.add_global_parameters(global_parameters)


#####################################################################
# Tests for input to .add_slide
#####################################################################

def validate(config, valid, match="Invalid value for "):

    default_config = {"content": ["A text", content_dir + "cat.jpg", content_dir + "chips.pdf"]}
    default_config.update(config)

    report = PowerPointReport()
    if valid:
        report.add_slide(**default_config)
    else:
        with pytest.raises((ValueError, TypeError, IndexError), match=match):
            report.add_slide(**default_config)


# ------------------------------------------------------------------- #
def test_invalid_params():
    """ Test that invalid parameters are caught """
    report = PowerPointReport()
    with pytest.raises(ValueError, match="Invalid parameter .+ given for slide"):
        report.add_slide(invalid_param="invalid")


# ------------------------------------------------------------------- #
@pytest.mark.parametrize("content, valid",
                         [(content_dir + "colored_animals/(.*)_blue.jpg", True),
                          (content_dir + "colored_animals/([*)_blue.jpg", False)])
def test_regex_input(caplog, content, valid):
    config = {"content": content}

    report = PowerPointReport()
    report.add_slide(**config)

    if valid:
        assert "WARNING" not in caplog.text  # check that no warning is written
    else:
        assert "Pattern is not a valid regex:" in caplog.text  # check that warning is written
        assert report._slides[0]._boxes[0].content == content


# ------------------------------------------------------------------- #
@pytest.mark.parametrize("content, valid",
                         [("A text", True),
                          (content_dir + "cat.jpg", True),
                          ([], True),
                          (1, True)])
def test_content_input(content, valid):
    config = {"content": content}
    validate(config, valid)


# ------------------------------------------------------------------- #
# grouped content
@pytest.mark.parametrize("grouped_content, valid",
                         [(["string", content_dir + "colored_animals/(.*)_blue.jpg", content_dir + "colored_animals/(.*)_red.jpg"], True),
                          ([content_dir + "colored_animals/(.*)_(.*).jpg"], False),  # two groups
                          (["no", "groups"], True),  # no groups, but valid
                          ("A text", False)])   # not a list
def test_grouped_content(grouped_content, valid):
    config = {"content": None, "grouped_content": grouped_content, "missing_file": "text"}
    validate(config, valid)


def test_grouped_content_warning(caplog):
    """ Test that a warning is written when no groups are found """
    config = {"grouped_content": [content_dir + "colored_animals/.*_blue.jpg"],
              "missing_file": "text"}

    report = PowerPointReport()
    report.add_slide(**config)
    assert "WARNING" in caplog.text
    assert "does not contain a capturing group" in caplog.text


@pytest.mark.parametrize("missing_file", ["raise", "empty", "skip", "text"])
def test_grouped_missing(caplog, missing_file):
    """ Test that a warning is written when only some groups are found """

    config = {"grouped_content": ["string",
                                  content_dir + "colored_animals/(.*)_red.jpg",
                                  content_dir + "colored_animals/(.*)_yellow.jpg",
                                  content_dir + "colored_animals/(.*)_blue.jpg"],
              "missing_file": missing_file}

    report = PowerPointReport(verbosity=2)

    if missing_file == "raise":
        with pytest.raises(FileNotFoundError, match=r"Missing file\(s\) for grouped content pattern"):
            report.add_slide(**config)
    elif missing_file == "empty":
        report.add_slide(**config)
        assert "Adding empty box." in caplog.text
    elif missing_file == "skip":
        report.add_slide(**config)
        assert "Skipping this element." in caplog.text
    elif missing_file == "text":
        report.add_slide(**config)
        assert "Adding this element as text." in caplog.text


@pytest.mark.parametrize("common", ["string",
                                    content_dir + "colored_animals/dog_blue.jpg",
                                    "non_existing_file.jpg"])
def test_grouped_missing_common(caplog, common):
    """ Test if non-grouped file is missing """

    config = {"grouped_content": [common,
                                  content_dir + "colored_animals/(.*)_red.jpg"],  # red animal is present in all groups
              "missing_file": "raise"}

    report = PowerPointReport()
    if common == "non_existing_file.jpg":
        with pytest.raises(FileNotFoundError, match=r"No files were found for the pattern"):
            report.add_slide(**config)
    else:
        report.add_slide(**config)  # no error

        if "dog_blue" in common:   # warning that file does not contain a group
            assert "does not contain a capturing group " in caplog.text


@pytest.mark.parametrize("empty_slide", ["keep", "skip"])
def test_grouped_missing_empty(caplog, empty_slide):

    config = {"grouped_content": ["string", "no_match(.)+"],
              "missing_file": "skip",
              "empty_slide": empty_slide}

    report = PowerPointReport()
    report.add_slide(**config)

    if empty_slide == "keep":
        assert "Adding slide without content." in caplog.text
        assert len(report._slides) == 1

    elif empty_slide == "skip":
        assert len(report._slides) == 0


# ------------------------------------------------------------------- #
# Set title (all types can be converted to str)
@pytest.mark.parametrize("title, valid",
                         [("A title", True),
                          (None, True),
                          (1, True),
                          (dict, True)])
def test_title_input(title, valid):
    config = {"title": title}
    validate(config, valid)


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
                                            ([["0", "1"], ["2", "3"]], True),
                                            ([[0, 0], [0, 0]], True),     # more content than numbers in layout; will limit content and write a warning
                                            ([[5, 5], [6, 6]], True),     # more indexes in layout than content; will write a warning
                                            ([["astring", 1], [1, 1]], False),  # right size but invalid type
                                            ([[-2, 1], [1, 1]], False),   # right size and type but -2 is not valid
                                            ("invalid", False),           # invalid string
                                            ([[0, 1, 2], [3, 4]], False)  # inconsistent number of columns
                                            ])
def test_content_layout(content, valid):
    """ Test that content layout is correctly validated """
    config = {"content_layout": content}
    validate(config, valid)


# ------------------------------------------------------------------- #
# content alignment
@pytest.mark.parametrize("content_alignment, valid",
                         [("left", True),
                          ("center", True),
                          ("right", True),
                          ("upper", True),  # upper center
                          ("lower", True),  # lower center
                          ("center right", True),
                          (["left", "center"], True),
                          ("upper wherever", False),
                          ("invalid", False),
                          (0, False)])
def test_content_alignment(content_alignment, valid):
    config = {"content_alignment": content_alignment}
    validate(config, valid)


# ------------------------------------------------------------------- #
# margins
@pytest.mark.parametrize("parameter", ["outer_margin", "inner_margin", "top_margin", "bottom_margin", "left_margin", "right_margin"])
@pytest.mark.parametrize("margins, valid",
                         [(0.1, True),
                          ("1", True),
                          (0, True),
                          ("1cm", False),
                          (-2, False)])
def test_margins_input(margins, valid, parameter):

    config = {parameter: margins}
    validate(config, valid)


# ------------------------------------------------------------------- #
# ratios
@pytest.mark.parametrize("parameter", ["width_ratios", "height_ratios"])
@pytest.mark.parametrize("ratios, valid",
                         [([1, 2], True),
                          ("1, 2", True),  # comma separated string can be converted to list
                          ([1, 2, 3, 5, 6], True),  # too many ratios
                          ([1], True),  # too few ratios
                          ([0, 1], False),  # ratio cannot be 0
                          ("1 2 3", False),  # string cannot be converted to list
                          ([], False),  # No ratios
                          ([0, -2], False),
                          (False, False),
                          (0, False)])
def test_ratios_input(ratios, valid, parameter):
    config = {parameter: ratios}
    validate(config, valid)


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
                          ("False", True),
                          (2, True),
                          ("2", True),
                          ("invalid", False),
                          (0, False),
                          ([], False)])
def test_split_input(split, valid):
    config = {"split": split}
    validate(config, valid)


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
    validate(config, valid)


# ------------------------------------------------------------------- #
# filename alignment
@pytest.mark.parametrize("filename_alignment, valid",
                         [("left", True),
                          ("center", True),
                          ("right", True),
                          ("RIGHT", True),
                          (["right", "left"], True),
                          ("center right", False),
                          ("invalid", False),
                          (0, False)])
def test_filename_alignment(filename_alignment, valid):
    config = {"filename_alignment": filename_alignment, "content_alignment": "left", "show_filename": True}  # bounds of filename is dependent on content_alignment (!=center is handled differently)
    validate(config, valid, "Invalid value for 'filename_alignment'")


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
    validate(config, valid)


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
    validate(config, valid, "Invalid value for 'remove_placeholders'")


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
    validate(config, valid, "Invalid value for 'fontsize'")


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
    validate(config, valid, "Invalid value for 'pdf_pages'")


@pytest.mark.parametrize("pdf_pages, valid", [("all", False),
                                              ([1, 2], False),
                                              (1, True)])
def test_pdf_pages_grouped(pdf_pages, valid):

    config = {"content": None, "grouped_content": [content_dir + "pdfs/multidogs_([0-9]).pdf"], "pdf_pages": pdf_pages}
    validate(config, valid, "Invalid value for ")


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
    validate(config, valid, "Invalid value for ")


# ------------------------------------------------------------------- #
# empty_slide
@pytest.mark.parametrize("empty_slide, valid",
                         [("keep", True),
                          ("skip", True),
                          ("invalid", False),
                          (False, False)])
def test_empty_slide(empty_slide, valid):
    config = {"empty_slide": empty_slide}
    validate(config, valid, "Invalid value for ")


# ------------------------------------------------------------------- #
# integers
@pytest.mark.parametrize("parameter", ["n_columns", "dpi"])
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
    validate(config, valid, match=f"Invalid value for '{parameter}' parameter")


# ------------------------------------------------------------------- #
# max_pixels
@pytest.mark.parametrize("value, valid",
                         [("1e3", True),
                          (4, True),
                          (1, False),
                          ("2", False),
                          (-1, False),
                          ("invalid", False),
                          (False, False)])
def test_max_pixels_input(value, valid):
    config = {"max_pixels": value}
    validate(config, valid, match="Invalid value for 'max_pixels' parameter")


# ------------------------------------------------------------------- #
# show_borders
@pytest.mark.parametrize("value, valid",
                         [(True, True),
                          (False, True),
                          (1, False),  # not bool
                          ("invalid", False)])
def test_show_borders(value, valid):
    """ Test that borders are correctly set """

    config = {"show_borders": value}
    validate(config, valid, match="Invalid value for 'show_borders' parameter")


# ------------------------------------------------------------------- #
# invalid combinations
@pytest.mark.parametrize("params", [{"content": ["text"], "grouped_content": ["text"]},    # both content and grouped_content given
                                    {"content": None, "split": True}])   # content has to given when split is True
def test_combinations(params):

    validate(params, False, "Invalid input combination")
