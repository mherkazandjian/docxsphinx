"""``docx2md`` / ``docx2rst`` command-line entry points.

Thin argparse layer over :func:`docxsphinx.reverse.engine.run_pandoc`.
Deliberately free of business logic — anything non-trivial belongs in
``engine.py`` so it can be unit-tested without a subprocess.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .engine import (
    OutputFormat,
    PandocConversionError,
    PandocNotFoundError,
    PandocVersionError,
    run_pandoc,
)

_MD_FLAVOURS: tuple[OutputFormat, ...] = (
    'gfm', 'commonmark', 'markdown', 'markdown_strict',
)


def _common_parser(
    prog: str, default_format: OutputFormat, format_choices: tuple[str, ...],
) -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog=prog,
        description='Convert a .docx file via pandoc and write the result '
                    'to stdout or a file.',
    )
    p.add_argument('input', type=Path, help='path to the .docx file')
    p.add_argument(
        '-o', '--output', type=Path, default=None,
        help='output file (default: stdout)',
    )
    p.add_argument(
        '--format', dest='out_format',
        choices=list(format_choices), default=default_format,
        help=f'pandoc writer to use (default: {default_format})',
    )
    p.add_argument(
        '--extract-media', type=Path, default=None,
        help='extract embedded images into this directory '
             '(default: alongside -o if given, otherwise images are dropped)',
    )
    return p


def _run(args: argparse.Namespace) -> int:
    """Dispatch one CLI invocation. Returns an exit code."""
    media_dir = args.extract_media
    if media_dir is None and args.output is not None:
        # Default behaviour when writing to a file: drop images into a
        # sibling directory named <basename>_media/.
        media_dir = args.output.parent / f'{args.output.stem}_media'

    try:
        text = run_pandoc(args.input, args.out_format, extract_media=media_dir)
    except PandocNotFoundError as exc:
        print(f'error: {exc}', file=sys.stderr)
        return 127
    except PandocVersionError as exc:
        print(f'error: {exc}', file=sys.stderr)
        return 2
    except PandocConversionError as exc:
        print(f'error: {exc}', file=sys.stderr)
        return exc.returncode or 1
    except FileNotFoundError as exc:
        print(f'error: {exc}', file=sys.stderr)
        return 2

    if args.output is None:
        sys.stdout.write(text)
    else:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text)
    return 0


def main_md(argv: list[str] | None = None) -> int:
    """Entry point for the ``docx2md`` console script."""
    parser = _common_parser(
        prog='docx2md',
        default_format='gfm',
        format_choices=_MD_FLAVOURS,
    )
    return _run(parser.parse_args(argv))


def main_rst(argv: list[str] | None = None) -> int:
    """Entry point for the ``docx2rst`` console script (RST always)."""
    parser = _common_parser(
        prog='docx2rst',
        default_format='rst',
        format_choices=('rst',),
    )
    return _run(parser.parse_args(argv))


if __name__ == '__main__':  # pragma: no cover
    sys.exit(main_md())
