"""Sphinx extension that registers the ``docx`` builder."""
from __future__ import annotations

from docxsphinx.builder import DocxBuilder


def setup(app):
    app.add_builder(DocxBuilder)
    app.add_config_value('docx_template', None, 'env')
    app.add_config_value('docx_debug_log', None, 'env')
    app.add_config_value('docx_documents', [], 'env')
    return {'parallel_read_safe': True, 'parallel_write_safe': True}
