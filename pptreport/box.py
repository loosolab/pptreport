import os
import re
import pkg_resources

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.text.layout import TextFitter

# For reading pictures
from PIL import Image


def split_string(string, length):
    """ Split a string into a list of strings of length 'length' """
    return [string[i:i + length] for i in range(0, len(string), length)]


def estimate_fontsize(txt_frame, min_size=6, max_size=18):
    """
    Resize text to fit the textbox.

    Parameters
    ----------
    txt_frame : pptx.text.text.TextFrame
        The text frame to be resized.
    min_size : int, default 6
        The minimum fontsize of the text.
    max_size : int, default 18
        The maximum fontsize of the text.

    Returns
    --------
    size : int
        The estimated fontsize of the text.
    """

    # Get the text across all runs
    text = ""
    for paragraph in txt_frame.paragraphs:
        for run in paragraph.runs:
            text += run.text

    # Get font
    font = pkg_resources.resource_filename("pptreport", "fonts/OpenSans-Regular.ttf")

    # Calculate best fontsize
    try:
        size = TextFitter.best_fit_font_size(text, txt_frame._extents, max_size, font)

    except TypeError as e:  # happens with long filenames, which cannot fit on one line

        # Try fitting by splitting long words; decrease length if TextFitter still fails
        max_word_len = 20
        while True:
            if max_word_len < 5:
                raise e
            try:
                words = text.split(" ")
                words = [split_string(word, max_word_len) for word in words]
                words = sum(words, [])  # flatten list
                text = " ".join(words)

                size = TextFitter.best_fit_font_size(text, txt_frame._extents, max_size, font)
                break  # success

            except TypeError:
                max_word_len = int(max_word_len / 2)  # decrease word length

    # the output of textfitter is None if the text does not fit; set text to smallest size
    if size is None:
        size = min_size

    # Make sure size is within bounds
    size = max(min_size, size)
    size = min(max_size, size)

    return size


def format_textframe(txt_frame, size=12, name="Calibri"):
    """ Set the fontsize of the text in the text frame. """

    for paragraph in txt_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)
            run.font.name = "Calibri"


class Box():
    """ A box is a constrained area of the slide which contains a single element e.g. text, a picture, a table, etc. """

    def __init__(self, slide, coordinates):
        """
        Initialize a box.

        Parameters
        ----------
        slide : pptx slide object
            The slide on which the box is located.
        coordinates : tuple
            Coordinates containing (left, top, width, height) of the box (in pptx units).
        """

        self.slide = slide
        self.logger = None

        # Bounds of the box
        self.left = int(coordinates[0])
        self.top = int(coordinates[1])
        self.width = int(coordinates[2])
        self.height = int(coordinates[3])

        # Initialize bounds of the content (can be smaller than the box)
        self.content = None
        self.content_left = self.left
        self.content_top = self.top
        self.content_width = self.width
        self.content_height = self.height

        self.border = None  # border object of the box

    def add_parameters(self, parameters):
        """ Add parameters from the slide """

        for key, value in parameters.items():
            setattr(self, key, value)

    def add_border(self):
        """ Adds a border shape of box to make debugging easier """

        if self.border is None:
            self.border = self.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.left, self.top, self.width, self.height)
            self.border.fill.background()
            self.border.line.color.rgb = RGBColor(0, 0, 0)  # black
            self.border.line.width = Pt(1)

    def remove_border(self):
        """ Removes the border shape """

        if self.border is not None:
            self.border._sp.getparent().remove(self.border._sp)
            self.border = None  # reset border

    @staticmethod
    def _get_content_type(content):
        """ Determine the type of content. """

        if isinstance(content, str):
            if os.path.isfile(content):
                # Find out the content of the file
                try:
                    with open(content) as f:
                        _ = f.read()
                    return "textfile"

                except UnicodeDecodeError:
                    return "image"
            else:
                return "text"
        elif content is None:
            return "empty"
        else:
            t = type(content)
            raise ValueError(f"Content of type '{t}' cannot be added to slide")

    def fill(self, content, box_index=0):
        """
        Fill the box with content. The function estimates type of content is given.

        Parameters
        ----------
        content : str
            The element to be added to the box.
        box_index : int
            The index of the box (used for accessing properties per box).
        """

        self.box_index = box_index
        self.content = content

        # Find out what type of content it is
        content_type = self._get_content_type(content)
        self.content_type = content_type

        if content_type == "image":
            full_height = self.height

            if self.show_filename is not False:
                # set height of filename to 1/10 of the textbox but at least 290000 (matches Calibri size 12) to ensure the text is still readable
                text_height = max(self.height * 0.1, 290000)
                text_top = self.top
                self.height = full_height - text_height
                self.top = self.top + text_height
            self.fill_image(content)

            if self.show_filename is not False:
                self.height = text_height
                self.top = text_top
                vertical, horizontal = self._get_content_alignment()
                if horizontal != "center":
                    self.left = self.picture.left
                    self.width = self.picture.width

                # Determine filename
                filename = self._filename

                if self.show_filename is True or self.show_filename == "filename":
                    filename = os.path.splitext(os.path.basename(filename))[0]  # basename without extension
                elif self.show_filename == "filename_ext":
                    filename = os.path.basename(filename)  # basename with extension
                elif self.show_filename == "filepath":
                    filename = os.path.splitext(filename)[0]  # filepath without extension
                elif self.show_filename == "filepath_ext":
                    filename = filename  # filepath with extension (original full path)
                elif self.show_filename == "path":
                    filename = os.path.dirname(filename)  # path without filename

                self.fill_text(filename, is_filename=True)

        elif content_type == "textfile":  # textfile can also contain markdown
            with open(content) as f:
                text = f.read()
            self.fill_text(text)

        elif content_type == "text":  # text can also contain markdown
            self.fill_text(content)

        elif content_type == "empty":
            return  # do nothing
        else:
            pass

        self.logger.debug(f"Box index {box_index} was filled with {content_type}")

    def fill_image(self, filename):
        """ Fill the box with an image. """

        # Find out the size of the image
        self._adjust_image_size(filename)
        self._adjust_image_position()  # adjust image position to middle of box

        # Add image
        self.logger.debug("Adding image to slide from file: " + filename)
        self.picture = self.slide.shapes.add_picture(filename, self.content_left, self.content_top, self.content_width, self.content_height)

    def _adjust_image_size(self, filename):
        """
        Adjust the size of the image to fit the box.

        Parameters
        ----------
        filename : str
            Path to the image file.
        """

        # Find out the size of the image
        im = Image.open(filename)
        im_width, im_height = im.size

        box_width = self.width
        box_height = self.height

        im_ratio = im_width / im_height  # >1 for landscape, <1 for portrait
        box_ratio = box_width / box_height

        # width is the limiting factor; height will be smaller than box_height
        if box_ratio < im_ratio:  # box is wider than image; height will be limiting
            self.content_width = box_width
            self.content_height = box_width * im_height / im_width  # maintain aspect ratio

        # height is the limiting factor; width will be smaller than box_width
        else:
            self.content_width = box_height * im_width / im_height  # maintain aspect ratio
            self.content_height = box_height

    def _get_content_alignment(self):
        """ Get the content alignment for this box. """

        self.logger.debug(f"Getting content alignment for box '{self.box_index}'. Input content alignment is '{self.content_alignment}'")

        if isinstance(self.content_alignment, str):  # if content alignment is a string, use it for all boxes
            this_alignment = self.content_alignment

        elif isinstance(self.content_alignment, list):  # if content alignment is a list, use the alignment for the current box
            if self.box_index > len(self.content_alignment) - 1:  # if box index is out of range, use default alignment
                this_alignment = "center"  # default alignment
            else:
                this_alignment = self.content_alignment[self.box_index]
        else:
            raise ValueError(f"Content alignment '{self.content_alignment}' is not valid. Valid content alignments are: str or list of str")

        # Check if current alignment is valid
        valid_alignments = ["left", "right", "center", "lower", "upper",
                            "lower left", "lower center", "lower right",
                            "upper left", "upper center", "upper right",
                            "center left", "center center", "center right"]

        if this_alignment.lower() not in valid_alignments:
            raise ValueError(f"Alignment '{self.content_alignment}' is not valid. Valid content alignments are: {valid_alignments}")

        # Expand into the structure "<vertical> <horizontal>"
        if this_alignment.lower() in ["left", "right", "center"]:
            this_alignment = "center " + this_alignment
        elif this_alignment.lower() in ["lower", "upper"]:
            this_alignment = this_alignment + " center"

        return this_alignment.split(" ")

    def _get_filename_alignment(self):
        """ Get the content alignment for this box. """

        self.logger.debug(f"Getting filename alignment for box '{self.box_index}'. Input filename alignment is '{self.filename_alignment}'")

        if isinstance(self.filename_alignment, str):  # if content alignment is a string, use it for all boxes
            this_alignment = self.filename_alignment

        elif isinstance(self.filename_alignment, list):  # if content alignment is a list, use the alignment for the current box
            if self.box_index > len(self.filename_alignment) - 1:  # if box index is out of range, use default alignment
                this_alignment = "center"  # default alignment
            else:
                this_alignment = self.filename_alignment[self.box_index]
        else:
            raise ValueError(f"Filename alignment '{self.filename_alignment}' is not valid. Valid filename alignments are: str or list of str")

        # Check if current alignment is valid
        valid_alignments = ["left", "right", "center"]

        if this_alignment.lower() not in valid_alignments:
            raise ValueError(f"Alignment '{self.filename_alignment}' is not valid. Valid filename alignments are: {valid_alignments}")

        return this_alignment

    def _adjust_image_position(self):
        """ Adjust the position of the image to be in the middle of the box. """

        # Get content alignment for this box
        vertical, horizontal = self._get_content_alignment()

        # Adjust image position vertically
        if vertical == "upper":
            self.content_top = self.top
        elif vertical == "lower":
            self.content_top = self.top + self.height - self.content_height
        elif vertical == "center":
            self.content_top = self.top + (self.height - self.content_height) / 2

        # Adjust image position horizontally
        if horizontal == "left":
            self.content_left = self.left
        elif horizontal == "right":
            self.content_left = self.left + self.width - self.content_width
        elif horizontal == "center":
            self.content_left = self.left + (self.width - self.content_width) / 2

    def _contains_md(self, text):
        """ Checks if a string contains any md sequences.
        """

        # https://chubakbidpaa.com/interesting/2021/09/28/regex-for-md.html
        md_regex = {"heading": r"(#{1,8}\s)(.*)",
                    "emphasis": r"(\*|\_)+(\S+)(\*|\_)+",
                    "links": r"(\[.*\])(\((http)(?:s)?(\:\/\/).*\))",
                    "images": r"(\!)(\[(?:.*)?\])\(.*(\.(jpg|png|gif|tiff|bmp))(?:(\s\"|\')(\w|\W|\d)+(\"|\'))?\)",
                    "uo-list": r"(^(\W{1})(\s)(.*)(?:$)?)+",
                    "io-list": r"(^(\d+\.)(\s)(.*)(?:$)?)+",
                    }

        for reg in md_regex.values():
            if re.search(reg, text):
                return True

        return False

    def _get_text_all_children(self, parent):
        """
        Small helper to improve robustness.
        Should normally be only one child per parent on this level.
        ! Use only internally for ast tree from mistune!
        """
        text = ""
        for c in parent["children"]:
            text += c["text"]
        return text

    def _fill_md(self, p, text):
        """
        Fills a paragraph p with basic markdown formatted text, like **Bold**, *italic* ,...
        Supported types:
        - Bold     **bold** / __bold__
        - Italic    *ital*  /  _ital_
        - Link     	[title](https://www.example.com)
        - Heading   #H1 / ## H2 / ...
        (- Image    Only partly - if alternative text is given it will be shown, image should be add via add_image())

        Parameters
        ----------
        p : <pptx.text.text._Paragraph>
            paragraph to add to
        text : str
            The text to be added to the box.
        """
        # mistune is only needed for md, only import if needed
        import mistune

        # render input as html.ast
        markdown = mistune.create_markdown(renderer="ast")

        # traverse the tree
        for string in text.split("\n"):
            for i, par in enumerate(markdown(string)):
                if par["type"] == "paragraph":

                    # Add newlines between paragraphs
                    if i > 0:
                        run = p.add_run()
                        run.text = "\n"

                    for child in par["children"]:  # children are the single md elements like bold, italic,...
                        # italic
                        if child["type"] == "emphasis":
                            run = p.add_run()
                            run.font.italic = True
                            run.text = self._get_text_all_children(child)
                        # bold
                        elif child["type"] == "strong":
                            run = p.add_run()
                            run.text = self._get_text_all_children(child)
                            run.font.bold = True
                        # link
                        elif child["type"] == "link":
                            run = p.add_run()
                            run.text = self._get_text_all_children(child)
                            hlink = run.hyperlink
                            hlink.address = child["link"]
                        # alternative text for images
                        elif child["type"] == "image":
                            try:
                                print("markdown images not supported. Trying alternative text.")
                                text = child["alt"]
                                run = p.add_run()
                                run.text = text
                            except KeyError:
                                print("No alternative text given. Skipping entry.")
                        # codespan
                        # elif child["type"]=="codespan":
                        # plain text & default case
                        else:
                            try:
                                text = child["text"]
                            except KeyError:
                                print("Unknown child type. Trying to append children's content as plain text.")
                                try:
                                    text = self._get_text_all_children(child)
                                except KeyError:
                                    print(f"Child type {child['type']} is not supported.")
                                    continue  # continue with next child
                            run = p.add_run()
                            run.text = text

                elif par["type"] == "newline":
                    run = p.add_run()
                    run.text = "\n\n"  # newline for previous line + newline for new paragraph

                elif par["type"] == "heading":
                    # implement heading (bold, bigger) ? Or add it only as plain txt
                    pass

                elif par["type"] == "list":
                    # implement list
                    pass
                else:
                    # will be handled in paragraph > codespan/link/block_code (they have duplicate entries in the tree)
                    pass

    def fill_text(self, text, is_filename=False):
        """
        Fill the box with text.

        Parameters
        ----------
        text : str
            The text to be added to the box.
        is_filename: bool, optional
            True if text contains a filename to be placed above image, False otherwise. Default: False
        """

        txt_box = self.slide.shapes.add_textbox(self.left, self.top, self.width, self.height)
        txt_frame = txt_box.text_frame
        txt_frame.word_wrap = True

        # Check if text contains markdown
        md = self._contains_md(text)

        # Place all text in one paragraph
        p = txt_frame.paragraphs[0]
        if md:
            self._fill_md(p=p, text=text)
        else:
            p.text = text

        # Try to fit text size to the box
        if self.fontsize is None:
            self.logger.debug("Estimating fontsize...")
            size = estimate_fontsize(txt_frame)
            self.logger.debug(f"Found: {size}")
        else:
            size = self.fontsize
        format_textframe(txt_frame, size=size)

        # Set alignment of text in textbox
        if is_filename:
            vertical = "lower"
            horizontal = self._get_filename_alignment()
        else:
            vertical, horizontal = self._get_content_alignment()

        if vertical == "upper":
            txt_frame.vertical_anchor = MSO_ANCHOR.TOP
        elif vertical == "lower":
            txt_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
        elif vertical == "center":
            txt_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        if horizontal == "left":
            p.alignment = PP_ALIGN.LEFT
        elif horizontal == "right":
            p.alignment = PP_ALIGN.RIGHT
        elif horizontal == "center":
            p.alignment = PP_ALIGN.CENTER