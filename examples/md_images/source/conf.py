"""Sphinx config for the Markdown-images docxsphinx example."""
project = 'md_images_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['myst_parser', 'docxsphinx']
source_suffix = {'.md': 'markdown'}
master_doc = 'index'

myst_enable_extensions = ['colon_fence', 'html_image', 'attrs_block']
exclude_patterns = ['_build']
