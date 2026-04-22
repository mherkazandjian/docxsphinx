"""Sphinx config for the Markdown-links docxsphinx example."""
project = 'md_links_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['myst_parser', 'docxsphinx']
source_suffix = {'.md': 'markdown'}
master_doc = 'index'

myst_enable_extensions = ['colon_fence', 'strikethrough']
# NOTE: MyST `linkify` (autodetection of bare URLs) requires the
# linkify-it-py optional dependency. This sample uses explicit
# [text](url) syntax to avoid pulling in that extra dependency.

exclude_patterns = ['_build']
