"""Tests for the reverse pipeline (docx → Markdown/RST via pandoc).

Mixed-tier file:

- Unit: :class:`subprocess.run` is monkey-patched to simulate pandoc's
  absence / failure / version mismatch. Fast, hermetic.
- Integration: real pandoc binary is invoked against a real docxsphinx-
  produced ``.docx`` (``examples/md_showcase/`` after a forward build).
  Asserts core content survives the round-trip.
- E2E: ``docx2md`` console script invoked as a subprocess. Confirms
  entry-point registration works end-to-end.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from docxsphinx.reverse import (
    PandocConversionError,
    PandocNotFoundError,
    PandocVersionError,
    docx_to_markdown,
    ensure_pandoc_available,
    pandoc_version,
    run_pandoc,
)
from docxsphinx.reverse.cli import main_md, main_rst

REPO_ROOT = Path(__file__).resolve().parent.parent
SHOWCASE_DOCX = REPO_ROOT / 'examples' / 'md_showcase' / 'build' / 'md_showcase_project-0.1.docx'


def _ensure_showcase_built() -> None:
    """Build md_showcase if its .docx isn't already present. Integration +
    e2e tests use it as their round-trip input."""
    if SHOWCASE_DOCX.is_file():
        return
    subprocess.run(
        [sys.executable, '-m', 'sphinx', '-b', 'docx', 'source', 'build'],
        cwd=REPO_ROOT / 'examples' / 'md_showcase',
        check=True, capture_output=True,
    )


# ---------------------------------------------------------------------------
# Unit — engine.py with mocked subprocess
# ---------------------------------------------------------------------------

@pytest.mark.unit
def test_pandoc_version_returns_none_when_missing() -> None:
    with patch('docxsphinx.reverse.engine.shutil.which', return_value=None):
        assert pandoc_version() is None


@pytest.mark.unit
def test_ensure_pandoc_available_raises_when_missing() -> None:
    with (
        patch('docxsphinx.reverse.engine.shutil.which', return_value=None),
        pytest.raises(PandocNotFoundError),
    ):
        ensure_pandoc_available()


@pytest.mark.unit
def test_ensure_pandoc_available_raises_on_old_version() -> None:
    class FakeResult:
        stdout = 'pandoc 1.19.2.4\nCompiled …\n'

    with patch('docxsphinx.reverse.engine.shutil.which', return_value='/usr/bin/pandoc'), patch(
        'docxsphinx.reverse.engine.subprocess.run',
        return_value=FakeResult(),
    ), pytest.raises(PandocVersionError) as exc_info:
        ensure_pandoc_available()
    assert 'too old' in str(exc_info.value)


@pytest.mark.unit
def test_run_pandoc_raises_missing_input(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError):
        run_pandoc(tmp_path / 'does-not-exist.docx')


@pytest.mark.unit
def test_run_pandoc_raises_conversion_error_on_nonzero_exit(tmp_path: Path) -> None:
    """Pandoc is present but returns non-zero — we surface stderr."""
    fake_docx = tmp_path / 'fake.docx'
    fake_docx.write_bytes(b'not a real docx')

    # First call is `pandoc --version` for the availability check — make it succeed.
    # Second call is the actual conversion — make it fail.
    class VersionResult:
        stdout = 'pandoc 3.1.11.1\n'
        returncode = 0

    class FailResult:
        stdout = ''
        stderr = 'pandoc: fake.docx: couldn\'t parse as docx\n'
        returncode = 1

    with patch(
        'docxsphinx.reverse.engine.subprocess.run',
        side_effect=[VersionResult(), FailResult()],
    ), patch(
        'docxsphinx.reverse.engine.shutil.which',
        return_value='/usr/bin/pandoc',
    ), pytest.raises(PandocConversionError) as exc_info:
        run_pandoc(fake_docx)
    assert exc_info.value.returncode == 1
    assert 'couldn\'t parse' in exc_info.value.stderr


# ---------------------------------------------------------------------------
# Integration — real pandoc + real docxsphinx-built docx
# ---------------------------------------------------------------------------

@pytest.mark.integration
def test_real_roundtrip_preserves_major_features() -> None:
    """Run md_showcase.docx back through pandoc and verify all major content
    types survive the trip."""
    _ensure_showcase_built()
    md = docx_to_markdown(SHOWCASE_DOCX)

    # Headings
    assert '# md_showcase' in md
    assert '## Inline formatting' in md
    assert '### Level three' in md

    # Inline formatting
    assert '**bold**' in md
    assert '*italic*' in md
    assert '~~strikethrough~~' in md

    # Lists
    assert '- alpha' in md or '-   alpha' in md

    # Code block survives as text. docxsphinx's Pygments-coloured output
    # doesn't round-trip back to a fenced code block via pandoc (the runs
    # look like styled prose, not a code block — noted in
    # docs/roundtrip-findings.md). Assert the content is present even if
    # the fencing isn't.
    assert 'def fib' in md or 'def greet' in md or 'Fibonacci' in md

    # Hyperlink
    assert 'https://www.sphinx-doc.org/' in md


@pytest.mark.integration
def test_docx_to_markdown_accepts_string_path() -> None:
    _ensure_showcase_built()
    md = docx_to_markdown(str(SHOWCASE_DOCX))
    assert md.startswith('# md_showcase'), md[:60]


@pytest.mark.integration
def test_extract_media_directory_receives_images(tmp_path: Path) -> None:
    _ensure_showcase_built()
    media = tmp_path / 'media'
    docx_to_markdown(SHOWCASE_DOCX, extract_media=media)
    # md_showcase embeds image1.png at least once — pandoc extracts it.
    extracted = list(media.rglob('*'))
    assert any(
        p.is_file() and p.suffix.lower() in {'.png', '.jpg', '.jpeg', '.gif'}
        for p in extracted
    ), extracted


# ---------------------------------------------------------------------------
# E2E — console script as a subprocess
# ---------------------------------------------------------------------------

@pytest.mark.e2e
def test_docx2md_cli_to_stdout() -> None:
    _ensure_showcase_built()
    result = subprocess.run(
        [sys.executable, '-m', 'docxsphinx.reverse.cli', str(SHOWCASE_DOCX)],
        capture_output=True, text=True, check=False,
    )
    assert result.returncode == 0, result.stderr
    assert '# md_showcase' in result.stdout


@pytest.mark.e2e
def test_docx2md_cli_to_file(tmp_path: Path) -> None:
    _ensure_showcase_built()
    out = tmp_path / 'showcase.md'
    rc = main_md([str(SHOWCASE_DOCX), '-o', str(out)])
    assert rc == 0
    assert out.is_file()
    text = out.read_text()
    assert '# md_showcase' in text
    # Default-media extraction writes images alongside the output file.
    media_dir = tmp_path / 'showcase_media'
    assert media_dir.is_dir()


@pytest.mark.e2e
def test_docx2md_cli_errors_on_missing_input() -> None:
    rc = main_md(['/nonexistent/path.docx'])
    assert rc == 2


@pytest.mark.e2e
def test_docx2rst_cli_produces_rst_headings(tmp_path: Path) -> None:
    _ensure_showcase_built()
    out = tmp_path / 'showcase.rst'
    rc = main_rst([str(SHOWCASE_DOCX), '-o', str(out)])
    assert rc == 0
    text = out.read_text()
    # RST uses ===/--- underlines instead of # prefixes for headings.
    assert '=' * 10 in text or '-' * 10 in text, text[:400]


# ---------------------------------------------------------------------------
# Full-loop roundtrip — md_showcase.md → docx → md → docx. Proves the whole
# toolchain loop closes: docxsphinx's output is pandoc-ingestible, and the
# markdown pandoc produces is docxsphinx-ingestible.
# ---------------------------------------------------------------------------

_SPHINX_CONF = """\
project = 'roundtrip'
version = '1'
release = '1'
language = 'en'
extensions = ['myst_parser', 'docxsphinx']
source_suffix = {'.md': 'markdown'}
master_doc = 'index'
# `html_image` lets MyST understand the <img> tags pandoc emits for
# docx images that carry explicit width/height attributes. Without it
# the tags become raw HTML (dropped by visit_raw) and images are lost
# on the second forward build.
myst_enable_extensions = [
    'colon_fence', 'deflist', 'strikethrough', 'tasklist',
    'attrs_block', 'html_image',
]
exclude_patterns = ['_build']
"""


def _count_elements(docx_path: Path) -> dict:
    """Return coarse structural counts from a .docx file's document.xml."""
    import zipfile

    from lxml import etree

    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    with zipfile.ZipFile(docx_path) as zf:
        root = etree.fromstring(zf.read('word/document.xml'))
    counts = {'p': 0, 'tbl': 0, 'hyperlink': 0, 'drawing': 0, 'bookmarkStart': 0}
    for el in root.iter():
        tag = el.tag
        if not isinstance(tag, str) or not tag.startswith(W):
            continue
        local = tag[len(W):]
        if local in counts:
            counts[local] += 1
    return counts


@pytest.mark.e2e
def test_showcase_full_roundtrip_loop_closes(tmp_path: Path) -> None:
    """Closing-the-loop test for the comprehensive example:

    ``md_showcase/index.md → docxsphinx → .docx → docx2md → .md
    → docxsphinx → .docx``

    The second docxsphinx invocation must exit zero and produce a valid
    ``.docx`` containing ``word/document.xml``. Any regression in either
    direction of the pipeline breaks this.
    """
    import shutil
    import zipfile

    _ensure_showcase_built()

    # Step 1: reverse the forward-built docx back to Markdown via docx2md.
    reverse_md = tmp_path / 'reverse.md'
    rc = main_md([str(SHOWCASE_DOCX), '-o', str(reverse_md)])
    assert rc == 0, 'docx2md on showcase failed'
    assert reverse_md.is_file()

    # Step 2: assemble a minimal Sphinx project around the reverse-md.
    src = tmp_path / 'src'
    build = tmp_path / 'build'
    src.mkdir()
    (src / 'conf.py').write_text(_SPHINX_CONF)
    shutil.copy(reverse_md, src / 'index.md')
    # docx2md extracted media alongside the .md — carry it over so images
    # resolve in the second forward build.
    media_dir = tmp_path / 'reverse_media'
    if media_dir.is_dir():
        shutil.copytree(media_dir, src / 'reverse_media')

    # Step 3: forward-build that Markdown via docxsphinx.
    result = subprocess.run(
        [sys.executable, '-m', 'sphinx', '-b', 'docx', '-q', str(src), str(build)],
        capture_output=True, text=True, check=False,
    )
    assert result.returncode == 0, (
        f'second forward build failed\nstdout:\n{result.stdout}\n'
        f'stderr:\n{result.stderr}'
    )

    # Step 4: the resulting .docx must exist and be a valid zip.
    outputs = list(build.glob('*.docx'))
    assert len(outputs) == 1, outputs
    round_docx = outputs[0]
    with zipfile.ZipFile(round_docx) as zf:
        names = zf.namelist()
        assert 'word/document.xml' in names
        assert zf.read('word/document.xml').startswith(b'<?xml'), (
            'document.xml does not start with XML declaration'
        )


@pytest.mark.e2e
def test_showcase_roundtrip_preserves_structural_counts(tmp_path: Path) -> None:
    """After the full loop, paragraph / table / bookmark counts should be
    in the same ballpark as the original. Catches catastrophic content
    loss (e.g. a visitor silently dropping whole sections) without being
    so strict that pandoc's legitimate flattening fails the test.

    Tolerance is generous on purpose — pandoc's reverse flattens
    admonition tables into blockquotes and Pygments-coloured code into
    prose, which reduces both table and paragraph counts on the second
    forward pass. We require the roundtripped output to retain at least
    60% of the original count for each metric.
    """
    import shutil

    _ensure_showcase_built()
    orig_counts = _count_elements(SHOWCASE_DOCX)

    reverse_md = tmp_path / 'reverse.md'
    assert main_md([str(SHOWCASE_DOCX), '-o', str(reverse_md)]) == 0

    src = tmp_path / 'src'
    build = tmp_path / 'build'
    src.mkdir()
    (src / 'conf.py').write_text(_SPHINX_CONF)
    shutil.copy(reverse_md, src / 'index.md')
    media_dir = tmp_path / 'reverse_media'
    if media_dir.is_dir():
        shutil.copytree(media_dir, src / 'reverse_media')
    result = subprocess.run(
        [sys.executable, '-m', 'sphinx', '-b', 'docx', '-q', str(src), str(build)],
        capture_output=True, text=True, check=False,
    )
    assert result.returncode == 0, result.stderr

    outputs = list(build.glob('*.docx'))
    round_counts = _count_elements(outputs[0])

    # Paragraph count should stay well within an order of magnitude.
    for key in ('p', 'bookmarkStart'):
        assert round_counts[key] >= 0.6 * orig_counts[key], (
            f'{key}: roundtrip {round_counts[key]} is less than 60% '
            f'of original {orig_counts[key]}'
        )
    # At least some hyperlinks and images should survive.
    assert round_counts['hyperlink'] > 0, 'all hyperlinks lost in roundtrip'
    assert round_counts['drawing'] > 0, 'all images lost in roundtrip'
