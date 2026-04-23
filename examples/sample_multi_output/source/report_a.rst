Report A — first output
=======================

This is the root of the first ``.docx`` produced by the ``docx``
builder when ``docx_documents`` routes it to ``ReportA.docx``.

Summary
-------

Report A demonstrates the most common use of ``docx_documents``: two
independent master docs that share a Sphinx project but emit separate
Word files.

Details
-------

- Each output lives in its own document tree, with its own heading
  hierarchy and bookmarks.
- The project-level ``docx_template`` can be overridden per entry
  (``None`` keeps the default).
