"""End-to-end integration tests.

Invokes ``sphinx-build -b docx`` as a subprocess against each example project,
asserts the build succeeded (exit code 0), and validates that the emitted
``.docx`` is a well-formed zip containing ``word/document.xml``. This catches
both sphinx-build failures (missed by the legacy test that discarded the
return code) and silently-corrupt docx output.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.e2e

EXAMPLES = [
    # RST samples (pre-existing)
    ('sample_1', 'example-0.1.docx'),
    ('sample_2', 'my_foo_project-0.0.0.docx'),
    ('sample_3', 'my_foo_project-0.0.0.docx'),
    ('sample_4', 'my_foo_project-0.0.0.docx'),
    ('sample_5', 'example-0.1.docx'),
    ('sample_6', 'example-0.1.docx'),
    # Markdown / MyST samples (Phase 3)
    ('md_basic', 'md_basic_project-0.1.docx'),
    ('md_lists', 'md_lists_project-0.1.docx'),
    ('md_tasklist', 'md_tasklist_project-0.1.docx'),
    ('md_tables', 'md_tables_project-0.1.docx'),
    ('md_code', 'md_code_project-0.1.docx'),
    ('md_images', 'md_images_project-0.1.docx'),
    ('md_links', 'md_links_project-0.1.docx'),
    ('md_admonitions', 'md_admonitions_project-0.1.docx'),
    ('md_mixed', 'md_mixed_project-0.1.docx'),
    ('md_showcase', 'md_showcase_project-0.1.docx'),
    ('md_math', 'md_math_project-0.1.docx'),
]


@pytest.mark.parametrize(('sample_name', 'expected_filename'), EXAMPLES)
def test_examples(
    examples_root: Path,
    sample_name: str,
    expected_filename: str,
) -> None:
    example_dir = examples_root / sample_name
    build_dir = example_dir / 'build'
    shutil.rmtree(build_dir, ignore_errors=True)

    result = subprocess.run(
        [sys.executable, '-m', 'sphinx', '-b', 'docx', 'source', 'build'],
        cwd=example_dir,
        capture_output=True,
        text=True,
        check=False,
    )

    assert result.returncode == 0, (
        f'sphinx-build failed for {sample_name} (exit {result.returncode}).\n'
        f'stdout:\n{result.stdout}\n'
        f'stderr:\n{result.stderr}'
    )

    output_path = build_dir / expected_filename
    assert output_path.is_file(), f'expected output missing: {output_path}'

    with zipfile.ZipFile(output_path) as zf:
        names = zf.namelist()
        assert 'word/document.xml' in names, (
            f'{output_path} is not a valid docx (missing word/document.xml); '
            f'entries: {names[:10]}'
        )
        document_xml = zf.read('word/document.xml')
        assert document_xml.startswith(b'<?xml'), (
            f'{output_path}:word/document.xml does not start with XML declaration'
        )
