import numpy as np
import os
from pptx.util import Cm
from pptreport.box import Box


class Slide():
    """ An internal class for creating slides. """

    def __init__(self, slide, parameters={}):

        self._slide = slide  # Slide object from python-pptx
        self._boxes = []     # Boxes in the slide
        self.logger = None

        parameters = self.format_parameters(parameters)
        self.add_parameters(parameters)

    @staticmethod
    def format_parameters(parameters):
        """ Checks and formats specific slide parameters to the correct type. """

        parameters = parameters.copy()  # Make a copy to not change the original dict

        # Format "n_columns" to int
        if "n_columns" in parameters:  # hasattr(self, "n_columns"):
            try:
                parameters["n_columns"] = int(parameters["n_columns"])
            except ValueError:
                raise ValueError(f"Could not convert 'n_columns' parameter to int. The given value is: '{parameters['n_columns']}'. Please use an integer.")

        # Format "split" to bool
        if "split" in parameters:
            split = parameters["split"]
            if isinstance(split, str):
                if split.lower() in ["true", "1", "t", "y", "yes"]:
                    parameters["split"] = True
                elif split.lower() in ["false", "0", "f", "n", "no"]:
                    parameters["split"] = False
                else:
                    raise ValueError(f"Could not convert 'split' parameter to bool. The given value is: '{split}'. Please use 'True' or 'False'.")

        # Format "show_filename" to bool
        if "show_filename" in parameters:
            show_filename = parameters["show_filename"]
            if isinstance(show_filename, str):
                if show_filename.lower() in ["true", "1", "t", "y", "yes"]:
                    parameters["show_filename"] = True
                elif show_filename.lower() in ["false", "0", "f", "n", "no"]:
                    parameters["show_filename"] = False
                else:
                    raise ValueError(f"Could not convert 'show_filename' parameter to bool. The given value is: '{show_filename}'. Please use 'True' or 'False'.")

        return parameters

    def add_parameters(self, parameters):
        """ Add parameters to the slide as internal variables. """

        for key in parameters:
            if key != "self":
                setattr(self, key, parameters[key])

    def get_config(self):
        """ Get the config dictionary for this slide. """

        config = self.__dict__.copy()  # Make a copy to not change the original dict
        for key in list(config):
            if key.startswith("_") or key == "logger":
                del config[key]

        return config

    def set_layout_matrix(self):
        """ Get the content layout matrix for the slide. """

        # Get variables from self
        layout = self.content_layout
        n_elements = len(self.content)
        n_columns = self.n_columns

        # Get layout matrix depending on "layout" variable
        if layout == "grid":
            n_columns = min(n_columns, n_elements)  # number of columns cannot be larger than number of elements
            n_rows = int(np.ceil(n_elements / n_columns))  # number of rows to fit elements
            n_total = n_rows * n_columns

            intarray = list(range(n_elements))
            intarray.extend([np.nan] * (n_total - n_elements))

            layout_matrix = np.array(intarray).reshape((n_rows, n_columns))

        elif layout == "vertical":
            layout_matrix = np.array(list(range(n_elements))).reshape((n_elements, 1))

        elif layout == "horizontal":
            layout_matrix = np.array(list(range(n_elements))).reshape((1, n_elements))

        else:
            try:
                layout_matrix = np.array(layout)
            except ValueError:
                raise ValueError(f"Could not convert 'layout' parameter to numpy array. The given value is: '{layout}'. Please use a valid layout matrix.")

        self._layout_matrix = layout_matrix

    # -------------------  Fill slide with content  ------------------- #
    def _fill_slide(self):
        """ Fill the slide with content from the internal variables """

        self.set_title()
        self.add_notes()

        # Fill boxes with content
        if len(self.content) > 0:
            self.set_layout_matrix()
            self.create_boxes()       # Create boxes based on layout
            self.fill_boxes()         # Fill boxes with content

    def set_title(self):
        """ Set the title of the slide. Requires self.title to be set. """

        if self.title is not None:

            if self._slide.shapes.title is None:
                self.logger.warning("Could not set title of slide. The slide does not have a title box.")
            else:
                self._slide.shapes.title.text = self.title

    def add_notes(self):
        """ Add notes to the slide. """

        if self.notes is not None:
            if isinstance(self.notes, list):
                notes_string = ''
                for s in self.notes:
                    if os.path.exists(s):
                        with open(s, 'r') as f:
                            notes_string += f'\n{f.read()}'
                    else:
                        notes_string += f'\n{s}'
                notes_string = notes_string.lstrip()  # remove leading newline

            elif os.path.exists(self.notes):
                with open(self.notes, "r") as f:
                    notes_string = f.read()

            else:
                notes_string = self.notes

            self._slide.notes_slide.notes_text_frame.text = notes_string

    def create_boxes(self):
        """ Create boxes for the slide dependent on the internal layout matrix. """

        layout_matrix = self._layout_matrix
        nrows, ncols = layout_matrix.shape

        # Establish left/right/top/bottom margins (in cm)
        left_margin = self.outer_margin if self.left_margin is None else self.left_margin
        right_margin = self.outer_margin if self.right_margin is None else self.right_margin
        top_margin = self.outer_margin if self.top_margin is None else self.top_margin
        bottom_margin = self.outer_margin if self.bottom_margin is None else self.bottom_margin

        # Convert margins from cm to pptx units
        left_margin_unit = Cm(left_margin)
        right_margin_unit = Cm(right_margin)
        top_margin_unit = Cm(top_margin)
        bottom_margin_unit = Cm(bottom_margin)
        inner_margin_unit = Cm(self.inner_margin)

        # Add to top margin based on size of title
        if self._slide.shapes.title.text != "":
            top_margin_unit = self._slide.shapes.title.top + self._slide.shapes.title.height + top_margin_unit
        # else:
        #    sp = self._slide.shapes.title.element
        #    sp.getparent().remove(sp) # remove title box if title is empty

        # How many columns and rows are there?
        n_rows, n_cols = layout_matrix.shape

        # Get total height and width of pictures
        total_width = self._slide_width - left_margin_unit - right_margin_unit - (n_cols - 1) * inner_margin_unit
        total_height = self._slide_height - top_margin_unit - bottom_margin_unit - (n_rows - 1) * inner_margin_unit

        # Get column widths and row heights
        if self.width_ratios is None:
            widths = (np.ones(ncols) / ncols) * total_width
        else:
            widths = np.array(self.width_ratios) / sum(self.width_ratios) * total_width

        if self.height_ratios is None:
            heights = (np.ones(nrows) / nrows) * total_height
        else:
            heights = np.array(self.height_ratios) / sum(self.height_ratios) * total_height

        # Box coordinates
        box_numbers = layout_matrix[~np.isnan(layout_matrix)].flatten()
        box_numbers = sorted(set(box_numbers))  # unique box numbers
        box_numbers = [box_number for box_number in box_numbers if box_number >= 0]  # remove negative numbers
        for i in box_numbers:

            # Get column and row number
            coordinates = np.argwhere(layout_matrix == i)

            # Get upper left corner of box
            row, col = coordinates[0]
            left = left_margin_unit + np.sum(widths[:col]) + col * inner_margin_unit
            top = top_margin_unit + np.sum(heights[:row]) + row * inner_margin_unit

            # Get total width and height of box (can span multiple columns and rows)
            width = 0
            height = 0
            rows = set(coordinates[:, 0])
            for row in rows:
                height += heights[row]
            height += (len(rows) - 1) * inner_margin_unit  # add inner margins between rows

            cols = set(coordinates[:, 1])
            for col in cols:
                width += widths[col]
            width += (len(cols) - 1) * inner_margin_unit  # add inner margins between columns

            #  Create box
            self.add_box((left, top, width, height))

    def add_box(self, coordinates):
        """
        Add a box to the slide.

        Parameters
        ----------
        coordinates : tuple
            Coordinates containing (left, top, width, height) of the box (in pptx units).
        """

        box = Box(self._slide, coordinates)
        box.logger = self.logger  # share logger with box

        # Add specific parameters to box
        keys = ["content_alignment", "show_filename", "filename_alignment"]
        parameters = {key: getattr(self, key) for key in keys}
        box.add_parameters(parameters)

        # Add box object to list
        self._boxes.append(box)

    def fill_boxes(self):
        """ Fill the boxes with the elements in self.content """

        for i, element in enumerate(self.content):
            self._boxes[i].fill(element, box_index=i)
