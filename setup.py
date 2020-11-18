# -*- coding: utf-8 -*-
import os
import sys
from setuptools import setup, find_packages
from setuptools.command.test import test as TestCommand

BASEDIR = os.path.dirname(os.path.abspath(__file__))

version = '1.0.0'
long_description = \
        open(os.path.join(BASEDIR, "src", "README.md")).read() + \
        open(os.path.join(BASEDIR, "src", "TODO.txt")).read()

classifiers = [
    "Development Status :: 2 - Pre-Alpha",
    #"Development Status :: 3 - Alpha",
    #"Development Status :: 4 - Beta",
    "Intended Audience :: System Administrators",
    "License :: OSI Approved :: MIT License",
    "Programming Language :: Python",
    "Topic :: Office/Business :: Office Suites",
    "Topic :: Software Development :: Documentation",
    "Topic :: Text Processing :: Markup",
]


def read_requirements(tests=False):
    """
    Read the requirements.txt file and return its content w/o pytest
    :param tests:
    :return:
    """
    pkgs = list(map(str.strip, open('requirements.txt').readlines()))

    if tests:
        retval = list(filter(lambda pkg: 'pytest' in pkg, pkgs))
    else:
        retval = list(filter(lambda pkg: 'pytest' not in pkg, pkgs))

    return list(filter(lambda x: x is not '', retval))

class PyTest(TestCommand):
    user_options = [('pytest-args=', 'a', "Arguments to pass to py.test")]

    def initialize_options(self):
        TestCommand.initialize_options(self)
        self.pytest_args = []

    def finalize_options(self):
        TestCommand.finalize_options(self)
        self.test_args = []
        self.test_suite = True

    def run_tests(self):
        # import here, cause outside the eggs aren't loaded
        import pytest
        errno = pytest.main(self.pytest_args)
        sys.exit(errno)


setup(
     name='docxsphinx',
     version=version,
     description='Sphinx docx builder extension.',
     long_description=long_description,
     classifiers=classifiers,
     keywords=['sphinx', 'extension', 'builder', 'docx', 'OpenXML'],
     author='Mher Kazandjian, Hugo Buddelmeijer',
     author_email='mherkazandjian at gmail dot com, blackshift at gmail dot com',
     url='https://github.com/mherkazandjian/docxsphinx',
     license='MIT',
     packages=find_packages('src'),
     package_dir={'': 'src'},
     install_requires=read_requirements(),
     extras_require=dict(test=read_requirements(tests=True)),
     tests_require=read_requirements(tests=True),
     cmdclass={'test': PyTest},
     zip_safe=False,
     entry_points={
        'sphinx.builders': [
            'docx=docxsphinx',
        ],
     }
)
