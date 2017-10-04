=================================
Welcome to Python's docx module
=================================

Make and edit docx in 200 lines of pure Python
===============================================

The module was created when I was looking for a Python support for MS Word .doc files on PyPI and Stackoverflow. Unfortunately, the only solutions I could find used:

1. COM automation
2. .net or Java
3. Automating OpenOffice or MS Office

For those of us who prefer something simpler, I made docx.

Making documents
=================

The docx module has the following features:

* Paragraphs
* Bullets
* Numbered lists
* Multiple levels of headings
* Tables
* Document Properties

Tables are just lists of lists, like this:

== == ==
A1 A2 A3
B1 B2 B3
C1 C2 C3
== == ==

Editing documents
==================

Thanks to the awesomeness of the lxml module, we can:

* Search and replace
* Extract plain text of document
* Add and delete items anywhere within the document

.. figure:: image1.png

    This is a test description


.. .. page-break::

Mathematical formulas
=====================

Pythagorean theorem :math:`a^2 + b^2 = c^2`.

Ideas? Questions? Want to contribute?
======================================

Email <python.docx@librelist.com>


.. toctree::
    :maxdepth: 1
    :hidden:

    index-ja
    restructuredtext

