"""Golden-fingerprint regression tests for every example project.

For each committed example, this test:

1. Cleans ``examples/<sample>/build/`` (forces a fresh build).
2. Runs ``sphinx-build -b docx`` via subprocess.
3. Extracts a deterministic text fingerprint (paragraph counts, style
   distribution, heading texts, table shapes, hyperlinks, bookmark
   names, etc. — see :mod:`tests._fingerprint`) from the resulting
   ``.docx``.
4. Diffs the fingerprint against ``tests/golden/<sample>.fp.txt``.

Legitimate output changes flow through a single gesture:

.. code-block::

    UPDATE_GOLDEN=1 pytest -m golden    # or: make golden-update

which rewrites the goldens and stages the diff for review in the next
PR.  Any regression that silently drops content (a visitor starts
``SkipNode``-ing, a helper stops emitting a bookmark, pygments output
collapses, etc.) surfaces as a visible text diff in CI.
"""
from __future__ import annotations

import difflib
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

# pytest prepends the test file's directory to sys.path, so these
# bare-module imports resolve without `tests/` needing to be a package.
from _fingerprint import fingerprint  # type: ignore[import-not-found]
from test_build_examples import EXAMPLES  # type: ignore[import-not-found]

pytestmark = [pytest.mark.e2e, pytest.mark.golden]

REPO_ROOT = Path(__file__).resolve().parent.parent
EXAMPLES_ROOT = REPO_ROOT / 'examples'
GOLDEN_DIR = Path(__file__).resolve().parent / 'golden'

UPDATE_GOLDEN = os.environ.get('UPDATE_GOLDEN') == '1'


def _build(sample_name: str, expected_filenames: tuple[str, ...]) -> list[Path]:
    """Force-build the example and return paths to every expected output."""
    example_dir = EXAMPLES_ROOT / sample_name
    build_dir = example_dir / 'build'
    shutil.rmtree(build_dir, ignore_errors=True)
    result = subprocess.run(
        [sys.executable, '-m', 'sphinx', '-b', 'docx', '-q', 'source', 'build'],
        cwd=example_dir, capture_output=True, text=True, check=False,
    )
    assert result.returncode == 0, (
        f'sphinx-build failed for {sample_name}:\n{result.stderr}'
    )
    paths = [build_dir / name for name in expected_filenames]
    for p in paths:
        assert p.is_file(), f'expected output missing: {p}'
    return paths


def _golden_name(sample_name: str, docx_name: str, single: bool) -> str:
    """Golden file stem: ``<sample>`` for single-output samples;
    ``<sample>__<docx-stem>`` when a sample produces multiple docs so
    each gets its own committed fingerprint."""
    if single:
        return sample_name
    return f'{sample_name}__{Path(docx_name).stem}'


@pytest.mark.parametrize(('sample_name', 'expected_filenames'), EXAMPLES)
def test_golden_fingerprint(
    sample_name: str, expected_filenames: tuple[str, ...],
) -> None:
    docxes = _build(sample_name, expected_filenames)
    single = len(docxes) == 1
    for docx, expected in zip(docxes, expected_filenames, strict=True):
        _compare_fingerprint(sample_name, docx, expected, single)


def _compare_fingerprint(
    sample_name: str, docx: Path, docx_name: str, single: bool,
) -> None:
    actual = fingerprint(docx)
    stem = _golden_name(sample_name, docx_name, single)
    golden_path = GOLDEN_DIR / f'{stem}.fp.txt'
    if UPDATE_GOLDEN or not golden_path.is_file():
        GOLDEN_DIR.mkdir(parents=True, exist_ok=True)
        golden_path.write_text(actual)
        if not UPDATE_GOLDEN:
            pytest.skip(
                f'wrote initial golden at {golden_path.relative_to(REPO_ROOT)}; '
                f're-run without UPDATE_GOLDEN to verify'
            )
        return

    expected = golden_path.read_text()
    if actual == expected:
        return

    diff = ''.join(difflib.unified_diff(
        expected.splitlines(keepends=True),
        actual.splitlines(keepends=True),
        fromfile=str(golden_path.relative_to(REPO_ROOT)),
        tofile=f'{sample_name} (fresh build)',
        n=3,
    ))
    pytest.fail(
        f'fingerprint mismatch for {sample_name} ({docx.name}).\n'
        f'If this change is intentional, regenerate goldens with:\n'
        f'    make golden-update\n'
        f'\ndiff:\n{diff}'
    )
