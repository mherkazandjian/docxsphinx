"""Sphinx config for the Markdown-lists docxsphinx example."""
project = 'md_lists_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['myst_parser', 'docxsphinx']
source_suffix = {'.md': 'markdown'}
master_doc = 'index'

myst_enable_extensions = ['colon_fence', 'deflist']
exclude_patterns = ['_build']
