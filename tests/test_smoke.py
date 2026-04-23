"""Smoke tests.

The fastest tier of the pyramid: does the package import, does the
Sphinx builder register, does ``setup(app)`` advertise the correct
parallel-safety metadata and config values?

These checks run in well under a second and should be the first signal
that something in packaging, imports, or extension wiring has broken.
"""
from __future__ import annotations

from unittest.mock import MagicMock

import pytest

pytestmark = pytest.mark.smoke


def test_package_imports() -> None:
    import docxsphinx  # noqa: F401


def test_builder_class_is_exposed() -> None:
    """``DocxBuilder`` must be importable from the top-level package."""
    from docxsphinx import DocxBuilder

    assert DocxBuilder.name == 'docx'
    assert DocxBuilder.format == 'docx'
    assert DocxBuilder.out_suffix == '.docx'


def test_entry_point_registered() -> None:
    """The ``sphinx.builders`` entry point must resolve to this package."""
    from importlib.metadata import entry_points

    points = [ep for ep in entry_points(group='sphinx.builders') if ep.name == 'docx']
    assert points, 'no sphinx.builders entry point named "docx" is registered'
    assert points[0].value == 'docxsphinx', points[0].value


def test_setup_returns_parallel_metadata_and_registers_config() -> None:
    """``setup(app)`` must register the builder, the project config
    values, and return parallel-safety metadata."""
    from docxsphinx import setup

    app = MagicMock()
    metadata = setup(app)

    assert metadata == {'parallel_read_safe': True, 'parallel_write_safe': True}

    # Builder registered exactly once.
    assert app.add_builder.call_count == 1
    (builder_arg,), _ = app.add_builder.call_args
    assert builder_arg.name == 'docx'

    # All config values are registered under their expected names.
    registered = {call.args[0] for call in app.add_config_value.call_args_list}
    assert registered == {
        'docx_template', 'docx_debug_log', 'docx_documents',
    }, registered
