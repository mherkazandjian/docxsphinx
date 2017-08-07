Introduction
============
This repository has been forked from

   https://bitbucket.org/shimizukawa/sphinxcontrib-docxbuilder

and some heavy modification have been done. The major changes are listed in
the release notes (`todo` add the release notes).

Installation
============
The latest bleeding edge version can be installed through

   pip install git+https://github.com/mherkazandjian/docxsphinx.git@master

Alternatively the repository can be cloned and installed locally

   git clone https://github.com/mherkazandjian/docxsphinx.git
   cd docxsphinx
   python -m pip install -r requirements.txt
   python setup.py install


After installation, ```make docx``` in a sphinx project  should work
(see next section)

Generating a `docx` document
============================
It is assumed that a sphinx project already is in place. At least one change
must be done to `conf.py` in-order to be able to generate a docx file.

The following line must be added to `conf.py`

    extensions = [
       'docxsphinx'
    ]


The sample projects are in the directory `examples`

  - REPO_ROOT/examples/sample_1 : default example (from the original repo)
  - REPO_ROOT/examples/sample_2 : example tested with `make docx`
  - REPO_ROOT/examples/sample_3 : example tested with `make docx` with a custom style


Word styles
===========

a custom word style file can be specified by adding

    # 'docx_template' need *.docx or *.dotx template file name. default is None.
    docx_template = 'template.docx'

to the end of `conf.py` (or anywhere in the file)

API
===
see also 

    REPO_ROOT/src/README.md  (outdated - but useful)
    REPO_ROOT/src/docxsphinx/docx/README.md
