"""pandoc subprocess wrapper used by the ``docx2md`` CLI and programmatic API.

Design intent: keep this module *purely* about invoking pandoc. No CLI
concerns, no argparse, no stdout writes. Callers receive either the
converted text as a string or a well-typed exception. Tests can monkey-
patch :func:`subprocess.run` to verify each failure mode.
"""
from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path
from typing import Literal

#: The lowest pandoc version we've tested against — GFM output was
#: introduced in pandoc 2.0. Older versions lack key writers we depend on.
MIN_PANDOC_VERSION: tuple[int, int] = (2, 0)

#: Output formats we officially expose through the reverse pipeline. These
#: map 1:1 to pandoc's own ``-t`` writer names.
OutputFormat = Literal['gfm', 'commonmark', 'markdown', 'markdown_strict', 'rst']

_PANDOC_VERSION_RE = re.compile(r'^pandoc\s+(\d+)\.(\d+)(?:\.(\d+))?', re.MULTILINE)


class ReverseError(RuntimeError):
    """Base class for all exceptions raised by this module."""


class PandocNotFoundError(ReverseError):
    """Raised when ``pandoc`` is not on ``PATH``. Gives a user-actionable hint."""

    def __init__(self) -> None:
        super().__init__(
            "pandoc executable not found on PATH. Install it — e.g. "
            "`apt install pandoc`, `brew install pandoc`, or `winget install "
            "--id JohnMacFarlane.Pandoc` — and re-run."
        )


class PandocVersionError(ReverseError):
    """Raised when the installed pandoc is older than :data:`MIN_PANDOC_VERSION`."""

    def __init__(self, found: tuple[int, int]) -> None:
        min_s = '.'.join(str(n) for n in MIN_PANDOC_VERSION)
        found_s = '.'.join(str(n) for n in found)
        super().__init__(
            f"pandoc {found_s} is too old; docxsphinx.reverse requires >= {min_s}"
        )


class PandocConversionError(ReverseError):
    """Raised when pandoc exits non-zero. ``stderr`` is attached for diagnostics."""

    def __init__(self, stderr: str, returncode: int) -> None:
        super().__init__(
            f"pandoc exited with code {returncode}. stderr:\n{stderr.rstrip()}"
        )
        self.stderr = stderr
        self.returncode = returncode


def pandoc_version() -> tuple[int, int, int] | None:
    """Return ``(major, minor, patch)`` of the installed pandoc, or ``None``
    if pandoc is not on ``PATH``. Raises nothing."""
    if shutil.which('pandoc') is None:
        return None
    try:
        result = subprocess.run(
            ['pandoc', '--version'],
            capture_output=True, text=True, check=False,
        )
    except OSError:
        return None
    match = _PANDOC_VERSION_RE.search(result.stdout)
    if not match:
        return None
    major, minor, patch = match.groups(default='0')
    return int(major), int(minor), int(patch)


def ensure_pandoc_available() -> None:
    """Assert pandoc is installed and at or above :data:`MIN_PANDOC_VERSION`.

    Raises :class:`PandocNotFoundError` or :class:`PandocVersionError`."""
    version = pandoc_version()
    if version is None:
        raise PandocNotFoundError
    if version[:2] < MIN_PANDOC_VERSION:
        raise PandocVersionError(version[:2])


def run_pandoc(
    input_path: Path,
    out_format: OutputFormat = 'gfm',
    *,
    extract_media: Path | None = None,
) -> str:
    """Convert ``input_path`` (a .docx file) to ``out_format`` via pandoc
    and return the result as a string.

    Parameters
    ----------
    input_path
        Path to the ``.docx`` file to convert.
    out_format
        One of :data:`OutputFormat`. Default ``'gfm'``.
    extract_media
        If given, pandoc extracts embedded media into this directory.
        Pass ``None`` to omit the ``--extract-media`` flag entirely (images
        are discarded from the text but not extracted to disk).

    Raises
    ------
    PandocNotFoundError
        pandoc is not on PATH.
    PandocVersionError
        pandoc is present but older than :data:`MIN_PANDOC_VERSION`.
    PandocConversionError
        pandoc ran but exited non-zero.
    FileNotFoundError
        ``input_path`` does not exist.
    """
    if not input_path.is_file():
        raise FileNotFoundError(f"docx input not found: {input_path}")
    ensure_pandoc_available()

    cmd = ['pandoc', '-f', 'docx', '-t', out_format, str(input_path.resolve())]
    cwd: Path | None = None
    if extract_media is not None:
        extract_media.mkdir(parents=True, exist_ok=True)
        # Run pandoc with CWD set to the media's parent so the
        # `--extract-media` value can be expressed relative to CWD. This
        # makes pandoc emit relative `<img src="…">` URIs in the output
        # markdown. If CWD and the media dir don't share a prefix,
        # pandoc falls back to absolute URIs, which break any downstream
        # MyST/sphinx-build rebuild of that markdown.
        cwd = extract_media.parent
        cmd.append(f'--extract-media={extract_media.name}')
    result = subprocess.run(
        cmd, capture_output=True, text=True, check=False, cwd=cwd,
    )
    if result.returncode != 0:
        raise PandocConversionError(result.stderr, result.returncode)
    return result.stdout
