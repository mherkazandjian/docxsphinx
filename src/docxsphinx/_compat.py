"""Version-compatibility shims for Sphinx, docutils, and python-docx.

All upstream imports that have moved between versions, or that are at risk of
being removed, route through this module. Import sites in ``builder.py`` and
``writer.py`` should import from here rather than from their original
locations, so there is one place to update when upstream breaks or moves an
API.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# sphinx.locale.__ — translation marker
# ---------------------------------------------------------------------------
try:
    from sphinx.locale import __
except ImportError:  # pragma: no cover - defensive; not known to have moved
    def __(message: str, *_args: object, **_kwargs: object) -> str:
        return message


# ---------------------------------------------------------------------------
# sphinx.util.osutil — filesystem helpers
# ---------------------------------------------------------------------------
try:
    from sphinx.util.osutil import ensuredir, os_path
except ImportError:  # pragma: no cover - defensive
    import os
    from pathlib import Path

    def ensuredir(path: str) -> None:
        Path(path).mkdir(parents=True, exist_ok=True)

    def os_path(path: str) -> str:
        return os.path.normpath(path)


# ---------------------------------------------------------------------------
# sphinx.util.nodes.inline_all_toctrees — used to fold every toctree into a
# single doctree so the docx builder can emit one document.
# ---------------------------------------------------------------------------
try:
    from sphinx.util.nodes import inline_all_toctrees
except ImportError as exc:  # pragma: no cover - defensive
    raise ImportError(
        "docxsphinx requires sphinx.util.nodes.inline_all_toctrees; your "
        "Sphinx release has removed it. Pin Sphinx to a version that still "
        "provides this API, or open an issue to track the replacement."
    ) from exc


# ---------------------------------------------------------------------------
# sphinx.util.console — ANSI colour helpers used only for build-time log
# prettification. Degrade to identity functions if they ever disappear.
# ---------------------------------------------------------------------------
try:
    from sphinx.util.console import bold, brown, darkgreen
except ImportError:  # pragma: no cover - defensive
    def _identity(text: str) -> str:
        return text

    bold = _identity
    brown = _identity
    darkgreen = _identity


# ---------------------------------------------------------------------------
# python-docx table cell class — currently exposed as a private name
# ``docx.table._Cell``. If it moves or becomes inaccessible, isinstance()
# checks against the stub below will always be False, which is the correct
# degradation (callers should then treat the location as a Document).
# ---------------------------------------------------------------------------
try:
    from docx.table import _Cell  # type: ignore[attr-defined]
except ImportError:  # pragma: no cover - defensive
    class _Cell:  # noqa: N801  (preserve upstream name)
        """Fallback stub; isinstance(obj, _Cell) is always False."""


__all__ = [
    "__",
    "_Cell",
    "bold",
    "brown",
    "darkgreen",
    "ensuredir",
    "inline_all_toctrees",
    "os_path",
]
