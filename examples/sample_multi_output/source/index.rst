sample_multi_output — many ``.docx`` from one project
=====================================================

This index is never built as an output itself — ``docx_documents`` in
``conf.py`` lists two explicit outputs (``ReportA.docx``,
``ReportB.docx``), each rooted at a different document. The index is
kept only so Sphinx's ``master_doc`` check is satisfied.

.. toctree::
   :hidden:

   report_a
   report_b
