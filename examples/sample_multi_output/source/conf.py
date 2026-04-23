"""Sphinx config demonstrating ``docx_documents`` — multiple ``.docx``
outputs from one Sphinx project, each rooted at a different master doc.

Mirrors the pattern Sphinx's built-in LaTeX builder uses via
``latex_documents``. The tuple format is:

    (startdoc, targetname, template, toctree_only)

- ``startdoc``: source docname (no suffix) — root of this output's
  toctree traversal.
- ``targetname``: output filename (``.docx`` suffix added if missing).
- ``template``: optional per-output template, overriding the project-
  level ``docx_template``. ``None`` → fall back to ``docx_template``.
- ``toctree_only``: when ``True``, the master's own prose content is
  stripped and only its toctree'd chapters are emitted. Useful when
  the master is a cover/index page only.
"""
project = 'sample_multi_output_project'
author = 'docxsphinx'
version = '0.1'
release = '0.1'
language = 'en'

extensions = ['docxsphinx']
master_doc = 'index'
exclude_patterns = ['_build']

docx_documents = [
    ('report_a', 'ReportA.docx', None, False),
    ('report_b', 'ReportB.docx', None, False),
]
