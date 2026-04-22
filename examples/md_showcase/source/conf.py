"""Sphinx config for the comprehensive Markdown showcase.

Exercises every docxsphinx feature from a single ``index.md``:
headings, inline formatting (bold / italic / strikethrough / sub / sup /
inline code), bullet + numbered + task + definition lists, GFM tables,
fenced code blocks with syntax highlighting, images with sizing and
alignment, external + internal hyperlinks (with inline formatting in link
text), admonitions (every type + generic), figure captions with
auto-numbering, and footnotes with inline formatting in the body.
"""
project = 'md_showcase_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['myst_parser', 'docxsphinx']

source_suffix = {'.md': 'markdown'}
master_doc = 'index'

myst_enable_extensions = [
    'colon_fence',      # :::{note} admonitions
    'deflist',          # definition lists
    'strikethrough',    # ~~text~~
    'tasklist',         # - [ ] / - [x]
    'attrs_block',      # {.class}/{image}-style attributes on blocks
    'attrs_inline',     # inline attrs for runs
    'html_image',       # <img> HTML for inline-in-paragraph images
    'substitution',     # {{ var }}
]
myst_enable_checkboxes = False

exclude_patterns = ['_build']
