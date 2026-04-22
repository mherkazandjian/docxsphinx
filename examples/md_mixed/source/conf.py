"""Sphinx config for the mixed RST+MD docxsphinx example.

Demonstrates that a single project can mix .rst and .md sources; the
docxsphinx builder inlines every toctree entry into one doctree and
emits a single .docx that contains the content of both.
"""
project = 'md_mixed_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['myst_parser', 'docxsphinx']
source_suffix = {
    '.rst': 'restructuredtext',
    '.md': 'markdown',
}
master_doc = 'index'

myst_enable_extensions = ['colon_fence']
exclude_patterns = ['_build']
