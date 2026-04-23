Report B — second output
========================

This is the root of the second ``.docx`` produced in the same
``sphinx-build`` invocation. Its content is completely independent of
``ReportA.docx`` — separate section tree, separate bookmark namespace
written into the output package.

Findings
--------

1. Multi-output builds are a single pass over the Sphinx environment.
2. Each entry in ``docx_documents`` recreates the writer, so nothing
   leaks from one output into the next.
3. Setting ``toctree_only=True`` on an entry drops the master's prose
   and emits only its toctree'd chapters — useful when the master is a
   cover page.
