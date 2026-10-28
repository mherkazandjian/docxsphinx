"""Sphinx config for the constructs docxsphinx 2.1 dropped silently.

Each block in ``index.rst`` carries a unique ``…MARK`` word so
``tests/test_dropped_content.py`` can assert that it reaches the ``.docx``.
"""
project = 'sample_directives_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['sphinx.ext.todo', 'docxsphinx']
master_doc = 'index'
exclude_patterns = ['_build']

todo_include_todos = True
