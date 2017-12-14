import os
from subprocess import Popen
import shlex
import shutil
import pytest

@pytest.mark.parametrize("example_dir, expected_docx_filename",
                         [
                             ('examples/sample_1', 'example-0.1.docx'),
                             ('examples/sample_2', 'my_foo_project-0.0.0.docx'),
                             ('examples/sample_3', 'my_foo_project-0.0.0.docx'),
                             ('examples/sample_4', 'my_foo_project-0.0.0.docx')
                         ])
def test_examples(example_dir, expected_docx_filename):
    build_dir = os.path.join(example_dir, 'build')
    shutil.rmtree(build_dir, ignore_errors=True)

    Popen(
        shlex.split(
            "sphinx-build -b docx source build"
        ),
        cwd=example_dir
    ).wait()
    assert os.path.isfile(os.path.join(build_dir, expected_docx_filename))
