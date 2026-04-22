"""Sphinx config for the GFM tasklist / checkbox docxsphinx example."""
project = 'md_tasklist_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = [
    'myst_parser',
    'docxsphinx',
]

source_suffix = {'.md': 'markdown'}
master_doc = 'index'

myst_enable_extensions = [
    'colon_fence',
    'tasklist',
]

# When True, MyST emits the tasklist checkbox as a clickable HTML <input>.
# We want the writer to have the information it needs to render a visible
# character in Word; set False so the class-attribute path is consistent.
myst_enable_checkboxes = False

exclude_patterns = ['_build']
