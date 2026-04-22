"""Sphinx config for the LaTeX math example.

Uses MyST's ``dollarmath`` and ``amsmath`` extensions so ``$...$`` and
``$$...$$`` become ``math`` / ``math_block`` nodes in the doctree, and
``\\begin{equation}...\\end{equation}`` environments are recognised as
display math. docxsphinx converts both via pandoc into Word-native OMML.

pandoc must be installed for math to render as OMML; without it the
writer falls back to emitting the raw LaTeX as a monospace run and
logs a warning.
"""
project = 'md_math_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['myst_parser', 'docxsphinx']
source_suffix = {'.md': 'markdown'}
master_doc = 'index'

myst_enable_extensions = [
    'colon_fence',
    'dollarmath',
    'amsmath',
]

exclude_patterns = ['_build']
