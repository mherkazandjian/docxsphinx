"""Sphinx config for the autodoc example — exercises sphinx.ext.autodoc
through docxsphinx so method signatures, parameter lists, and docstring
field lists (``:param:`` / ``:returns:`` / ``:raises:``) render into Word."""
import os
import sys

sys.path.insert(0, os.path.abspath('.'))

project = 'sample_autodoc_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = [
    'sphinx.ext.autodoc',
    'docxsphinx',
]

master_doc = 'index'
exclude_patterns = ['_build']

autodoc_member_order = 'bysource'
