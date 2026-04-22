"""Reverse pipeline — convert ``.docx`` files to Markdown (and later RST) via pandoc.

Public API::

    from docxsphinx.reverse import docx_to_markdown, docx_to_rst
    md = docx_to_markdown('report.docx')

CLI::

    docx2md report.docx -o report.md

Pandoc must be installed and on ``PATH``. See
:class:`~docxsphinx.reverse.engine.PandocNotFoundError` for the error raised
when it isn't.
"""
from __future__ import annotations

from pathlib import Path

from .engine import (
    MIN_PANDOC_VERSION,
    PandocConversionError,
    PandocNotFoundError,
    PandocVersionError,
    ReverseError,
    ensure_pandoc_available,
    pandoc_version,
    run_pandoc,
)


def docx_to_markdown(
    path: str | Path,
    *,
    flavour: str = 'gfm',
    extract_media: str | Path | None = None,
) -> str:
    """Convert a ``.docx`` file to Markdown and return the result.

    ``flavour`` selects the pandoc writer (``gfm``, ``commonmark``,
    ``markdown``, ``markdown_strict``). ``extract_media``, if given,
    tells pandoc to extract embedded images into that directory.
    """
    media = Path(extract_media) if extract_media is not None else None
    return run_pandoc(Path(path), out_format=flavour, extract_media=media)  # type: ignore[arg-type]


def docx_to_rst(
    path: str | Path,
    *,
    extract_media: str | Path | None = None,
) -> str:
    """Convert a ``.docx`` file to reStructuredText and return the result."""
    media = Path(extract_media) if extract_media is not None else None
    return run_pandoc(Path(path), out_format='rst', extract_media=media)


__all__ = [
    'MIN_PANDOC_VERSION',
    'PandocConversionError',
    'PandocNotFoundError',
    'PandocVersionError',
    'ReverseError',
    'docx_to_markdown',
    'docx_to_rst',
    'ensure_pandoc_available',
    'pandoc_version',
    'run_pandoc',
]
