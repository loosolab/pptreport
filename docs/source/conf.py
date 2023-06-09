# Configuration file for the Sphinx documentation builder.
#
# This file only contains a selection of the most common options. For a full
# list see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Path setup --------------------------------------------------------------

# If extensions (or modules to document with autodoc) are in another directory,
# add these directories to sys.path here. If the directory is relative to the
# documentation root, use os.path.abspath to make it absolute, like shown here.
#

import os
import sys

cwd = os.getcwd()

sys.path.insert(0, cwd)
sys.path.insert(0, os.path.abspath('../..'))


# -- Project information -----------------------------------------------------

project = 'pptreport'
copyright = '2023, Loosolab'
author = 'Loosolab'


# ----------------------------------------------------------------------------

# Run all examples and build the .rst file with the output

import build_examples
build_examples.main()  # run function from build_examples.py

# -- General configuration ---------------------------------------------------

# Add any Sphinx extension module names here, as strings. They can be
# extensions coming with Sphinx (named 'sphinx.ext.*') or your custom
# ones.
extensions = ['sphinxcontrib.images',
              'sphinx.ext.autodoc',
              'sphinx.ext.napoleon',
              'myst_parser',
              # 'sphinx.ext.viewcode',
              # 'sphinx.ext.intersphinx',
              # "nbsphinx",
              # "nbsphinx_link",
              ]

napoleon_numpy_docstring = True
autodoc_member_order = 'bysource'

# Add any paths that contain templates here, relative to this directory.
templates_path = ['_templates']

# List of patterns, relative to source directory, that match files and
# directories to ignore when looking for source files.
# This pattern also affects html_static_path and html_extra_path.
exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store']

images_config = {
    "default_image_width": "15%"
}
html_static_path = ['_static']
html_css_files = ["thumbnails.css"]

# -- Options for HTML output -------------------------------------------------

# The theme to use for HTML and HTML Help pages.  See the documentation for
# a list of builtin themes.
#
# html_theme = 'alabaster'
html_theme = 'sphinx_rtd_theme'
