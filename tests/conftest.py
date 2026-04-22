"""Shared pytest fixtures for the docxsphinx test suite."""
from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from types import SimpleNamespace

import pytest
from docutils.core import publish_doctree
from docx import Document
from docx.document import Document as DocumentType

pytest_plugins = ('sphinx.testing.fixtures',)

REPO_ROOT = Path(__file__).resolve().parent.parent
EXAMPLES_ROOT = REPO_ROOT / 'examples'


@pytest.fixture
def repo_root() -> Path:
    """Absolute path to the repository root, regardless of pytest invocation cwd."""
    return REPO_ROOT


@pytest.fixture
def examples_root() -> Path:
    """Absolute path to the examples/ directory."""
    return EXAMPLES_ROOT


@pytest.fixture
def rootdir() -> Path:
    """Root for sphinx.testing.fixtures — points at examples/ for in-memory builds."""
    return EXAMPLES_ROOT


@pytest.fixture
def fake_builder() -> SimpleNamespace:
    """Minimal ``Builder`` stand-in suitable for driving ``DocxTranslator``."""
    return SimpleNamespace(
        env=SimpleNamespace(srcdir='.'),
        config={'docx_template': None, 'docx_debug_log': None},
    )


@pytest.fixture
def translator_factory(fake_builder: SimpleNamespace) -> Callable[[str, str], DocumentType]:
    """Factory: parse a source string with the chosen parser and run it through the visitor.

    Call as ``translator_factory(src, 'rst')`` or ``translator_factory(src, 'md')``
    and receive a python-docx ``Document``. Lets integration tests exercise the
    full parse→doctree→visit→Document pipeline without touching disk or subprocess.
    """
    from docxsphinx.writer import DocxTranslator

    def _make(src: str, parser: str = 'rst') -> DocumentType:
        if parser == 'rst':
            doctree = publish_doctree(src)
        elif parser == 'md':
            from myst_parser.parsers.docutils_ import Parser as MystParser
            doctree = publish_doctree(
                src,
                parser=MystParser(),
                settings_overrides={
                    'myst_enable_extensions': ['tasklist', 'colon_fence', 'deflist'],
                    'myst_enable_checkboxes': False,
                },
            )
        else:
            raise ValueError(f'unknown parser {parser!r}; expected "rst" or "md"')
        container = Document()
        translator = DocxTranslator(doctree, fake_builder, container)
        doctree.walkabout(translator)
        return container

    return _make
