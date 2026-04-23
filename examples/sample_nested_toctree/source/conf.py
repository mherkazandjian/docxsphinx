"""Sphinx config for the nested-toctree heading-level regression example.

Reproduces issue #53: a master document that pulls a second document
via ``toctree``, which in turn pulls a third. Each included document's
top-level section should render one Word-heading level deeper than
the section that contains the toctree — matching Sphinx's ``singlehtml``
behaviour (<h1>, <h2>, <h3>). Before the fix, ``visit_start_of_file``
reset ``sectionlevel`` to zero on every file boundary, collapsing every
included heading to Heading 1.
"""
project = 'sample_nested_toctree_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['docxsphinx']
master_doc = 'index'
exclude_patterns = ['_build']
