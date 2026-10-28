Formerly dropped constructs
===========================

docxsphinx 2.1 silently dropped every construct on this page, content and
all. Each one carries a unique ``…MARK`` word so the test suite can check
that it reaches the generated ``.docx``.

.. contents:: On this page
   :local:

Docutils blocks
---------------

.. topic:: Topic title

   topicMARK body text.

.. sidebar:: Sidebar title
   :subtitle: Sidebar subtitle

   sidebarMARK body text.

.. rubric:: rubricMARK heading

| lineblockMARK first line with **bold** text
| second line
|     nested line

>>> doctestMARK = 1
>>> doctestMARK
1

.. container:: custom-class

   containerMARK paragraph.

A block quote with an attribution:

    quoteMARK text.

    -- attributionMARK Author

Sphinx directives
-----------------

.. versionadded:: 1.2
   versionaddedMARK body text.

.. versionchanged:: 1.3
   versionchangedMARK body text.

.. deprecated:: 2.0

.. todo:: todoMARK item.

.. figure:: image1.png
   :width: 4cm

   Figure caption.

   legendMARK paragraph under the caption.

.. code-block:: python
   :caption: codecaptionMARK
   :name: code-sample

   print("codeMARK")

See :ref:`code-sample` for the captioned block and :token:`expr` for the
grammar below.

.. productionlist::
   stmt: `expr` ";"
   expr: "x" | "y"

.. acks::

   * acksMARK Alice
   * Bob
