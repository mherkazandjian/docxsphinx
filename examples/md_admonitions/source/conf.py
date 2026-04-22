"""Sphinx config for the Markdown-admonitions docxsphinx example.

NOTE: until Phase 2.1 of the modernization plan lands, admonition content
is silently dropped by the writer (every ``visit_<type>`` admonition
method raises ``SkipNode``). The generated .docx for this sample will be
structurally valid but nearly empty. The sample exists so the Phase 2
work has a ready-made integration target.
"""
project = 'md_admonitions_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['myst_parser', 'docxsphinx']
source_suffix = {'.md': 'markdown'}
master_doc = 'index'

myst_enable_extensions = ['colon_fence', 'attrs_block']
exclude_patterns = ['_build']
