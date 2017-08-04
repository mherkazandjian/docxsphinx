Sphinx docx builder extension generate single docx file from Sphinx document
source. This extension use python-docx module (included) for the docx file
generation.

Features
========

* This extension work on Multi-platform (not need OpenOffice or MS Word).
* Usable sphinx syntax and directives:
    * heading line output
    * paragraph output (standard body text)
    * image and figure directive output
    * bullet-list and numbered-list output
    * table output (restrictive)
* You can use dotx/docx file for style template.

Currently, many directives and indented block are not work correctly, yet.

Setup
=====

Make environment by easy_install
---------------------------------

Not yet.

Make environment by buildout
-----------------------------

'hg clone' or download sphinxcontrib-docxbuilder archive from 'get source'
menu at http://bitbucket.org/shimizukawa/sphinxcontrib-docxbuilder ::

    $ cd /path/to/sphinxcontrib-docxbuilder
    $ python bootstrap.py -d init
    $ bin/buildout


run example
------------

for example sphinx-docx building, simply run below::

    $ bin/example
    ...
    Saved new file to: examples/example-0.1.docx


Usage
=====

Set 'sphinxcontrib-docxbuilder' to 'extensions' line of target sphinx source
conf.py::

    extensions = ['sphinxcontrib-docxbuilder']

If you want to use 'dotx' template, put 'template.dotx' file (that dotx's style
names are need to modifiy with `python-docx style name spec`_ ) into source
directory and write below spec in conf.py::

    docx_template = 'template.dotx'

Execute sphinx-build with below option::

    $ bin/sphinx-build -b docx [input-dir] [output-dir]
    $ ls [output-dir]
    output.docx


python-docx style name spec
============================

`dotx` is a template file was created with Word 2007 or later.
You can use `dotx`, but that need static-named style names.
`Style Name` such as 'Heading1' is constructed across python-docx module
by a use to specify the displaying of document data.
You must set style names by below names on Word 2007::

* Normal
* Heading1
* Heading2
* Heading3
* Heading4
* Heading5
* ListBullet
* ListNumber
* TableNormal

.. below names are not implemented by need at future.
.. * Title
.. * SubTitle
.. * NoList
.. * Strong
.. * Emphasis
.. * NoSpacing
.. * BlockQuote
.. * LiteralBlock
.. * BookTitle


License
========
Licensed under the `MIT license <http://www.opensource.org/licenses/mit-license.php>`_ .
See the LICENSE file for specific terms.


Requirements
============

* Python 2.6 or later (not support 3.x)
* `python-docx <http://github.com/mikemaccana/python-docx>`_
  (not released, but included), Thanks Mike MacCana.
* setuptools or distriubte.

History
=======

0.0.1 (unreleased)
--------------------
* Not released.


