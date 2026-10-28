"""End-to-end checks for content docxsphinx used to drop or crash on.

``examples/sample_directives`` holds every construct 2.1 skipped silently,
each tagged with a unique ``…MARK`` word. The first group of tests builds it
with ``sphinx-build`` and checks every word reaches ``word/document.xml``.

The second group builds a throwaway project in ``tmp_path`` to check that an
unsupported image and an unknown node type each produce one ``docx.*``
warning and a finished document, instead of a crash or a silent gap.
"""
from __future__ import annotations

import re
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.e2e

REPO_ROOT = Path(__file__).resolve().parent.parent
EXAMPLE_SRC = REPO_ROOT / 'examples' / 'sample_directives' / 'source'


def _sphinx_build(srcdir: Path, outdir: Path, *extra: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, '-m', 'sphinx', '-b', 'docx', '-q', '-E', *extra,
         str(srcdir), str(outdir)],
        capture_output=True, text=True, check=False,
    )


def _document_xml(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path) as zf:
        return zf.read('word/document.xml').decode()


@pytest.fixture(scope='module')
def directives_build(tmp_path_factory) -> tuple[str, str]:
    """Build the example once; return ``(document.xml, stderr)``."""
    out = tmp_path_factory.mktemp('sample_directives')
    result = _sphinx_build(EXAMPLE_SRC, out)
    assert result.returncode == 0, result.stderr
    return _document_xml(out / 'sample_directives_project-0.1.docx'), result.stderr


@pytest.mark.parametrize('marker', [
    'topic', 'sidebar', 'rubric', 'lineblock', 'doctest', 'container', 'quote',
    'attribution', 'versionadded', 'versionchanged', 'todo', 'legend',
    'codecaption', 'code', 'acks',
])
def test_formerly_dropped_construct_reaches_docx(
    directives_build: tuple[str, str], marker: str,
) -> None:
    xml, _ = directives_build
    assert f'{marker}MARK' in xml


def test_version_directive_labels_are_kept(directives_build: tuple[str, str]) -> None:
    """Sphinx injects "Added in version X" / "Deprecated since version X" into
    the body; a body-less ``deprecated`` is nothing *but* that label."""
    xml, _ = directives_build
    for version in ('version 1.2', 'version 1.3', 'version 2.0'):
        assert version in xml, version


def test_only_the_real_figure_gets_a_figure_number(directives_build: tuple[str, str]) -> None:
    """The captioned code-block's caption sits in a container now rendered
    in place; it must not be numbered as a figure."""
    xml, _ = directives_build
    assert xml.count('SEQ Figure') == 1


def test_internal_links_all_land_on_bookmarks(directives_build: tuple[str, str]) -> None:
    """``.. contents::`` entries, ``:ref:`` to the captioned code block and
    ``:token:`` into the grammar all need a matching bookmark."""
    xml, _ = directives_build
    anchors = set(re.findall(r'w:anchor="([^"]+)"', xml))
    bookmarks = set(re.findall(r'w:bookmarkStart w:id="\d+" w:name="([^"]+)"', xml))
    assert {'code-sample', 'grammar-token-expr', 'sphinx-directives'} <= anchors, anchors
    assert anchors <= bookmarks, anchors - bookmarks


def test_example_builds_without_docx_warnings(directives_build: tuple[str, str]) -> None:
    _, stderr = directives_build
    assert '[docx.' not in stderr, stderr


_MYSTERY_EXT = '''\
"""A directive whose node type docxsphinx has no visitor for."""
from docutils import nodes
from docutils.parsers.rst import Directive


class mystery(nodes.General, nodes.Element):
    pass


class Mystery(Directive):
    has_content = True

    def run(self):
        node = mystery()
        self.state.nested_parse(self.content, self.content_offset, node)
        return [node]


def setup(app):
    app.add_node(mystery)
    app.add_directive('mystery', Mystery)
'''

_PROBE_INDEX = '''\
Probe
=====

.. image:: pic.svg

.. mystery::

   mysteryMARK one.

.. mystery::

   mysteryMARK two.

After the mystery, afterMARK.
'''


def _probe_project(tmp_path: Path) -> Path:
    """A project with an SVG image and two nodes of an unknown type."""
    src = tmp_path / 'source'
    src.mkdir()
    (src / 'conf.py').write_text(
        "import os, sys\n"
        "sys.path.insert(0, os.path.abspath('.'))\n"
        "project = 'probe'\nversion = '0'\n"
        "extensions = ['mystery_ext', 'docxsphinx']\n"
    )
    (src / 'mystery_ext.py').write_text(_MYSTERY_EXT)
    (src / 'pic.svg').write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>'
    )
    (src / 'index.rst').write_text(_PROBE_INDEX)
    return src


def test_svg_and_unknown_node_warn_instead_of_failing(tmp_path: Path) -> None:
    """Before: the SVG aborted the build (UnrecognizedImageError) and the
    unknown node vanished without a word. Now: one warning each, exit 0, and
    the unknown node's text is in the document."""
    out = tmp_path / 'out'
    result = _sphinx_build(_probe_project(tmp_path), out)
    assert result.returncode == 0, result.stderr
    assert result.stderr.count('[docx.image]') == 1, result.stderr
    assert result.stderr.count('[docx.unknown_node]') == 1, result.stderr
    # The extension node has no line number; the warning borrows its body's.
    assert re.search(r'index\.rst:\d+: WARNING: docxsphinx has no renderer', result.stderr), (
        result.stderr
    )
    xml = _document_xml(out / 'probe-0.docx')
    assert xml.count('mysteryMARK') == 2
    assert 'afterMARK' in xml


def test_docx_warnings_honour_suppress_warnings(tmp_path: Path) -> None:
    out = tmp_path / 'out'
    result = _sphinx_build(
        _probe_project(tmp_path), out,
        '-D', 'suppress_warnings=docx.image,docx.unknown_node',
    )
    assert result.returncode == 0, result.stderr
    assert '[docx.' not in result.stderr, result.stderr
