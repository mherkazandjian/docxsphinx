# docxsphinx

[![CI](https://github.com/mherkazandjian/docxsphinx/actions/workflows/ci.yml/badge.svg)](https://github.com/mherkazandjian/docxsphinx/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/docxsphinx.svg)](https://pypi.org/project/docxsphinx/)
[![Python](https://img.shields.io/pypi/pyversions/docxsphinx.svg)](https://pypi.org/project/docxsphinx/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A Sphinx builder extension that emits **Microsoft Word (`.docx`)** from
Sphinx documentation sources. Accepts both reStructuredText and Markdown
(via [myst-parser](https://myst-parser.readthedocs.io/)) input — including
GFM task-list checkboxes, rendered as `☐` / `☒` glyphs in the Word output.

Forked from [shimizukawa/sphinxcontrib-docxbuilder](https://bitbucket.org/shimizukawa/sphinxcontrib-docxbuilder)
with substantial modification. See [`CHANGELOG.md`](CHANGELOG.md) for release notes.

## Installation

From PyPI:

```
pip install docxsphinx
```

From source:

```
pip install git+https://github.com/mherkazandjian/docxsphinx.git@master
```

## Usage

In your Sphinx project's `conf.py`:

```python
extensions = ['docxsphinx']
```

Then invoke the builder:

```
sphinx-build -b docx source build
```

The builder emits a single `.docx` named `{project}-{version}.docx` inside
the build directory (`project` and `version` come from `conf.py`).

### Markdown input

Add `myst-parser` and register both parsers:

```python
extensions = ['myst_parser', 'docxsphinx']
source_suffix = {
    '.rst': 'restructuredtext',
    '.md':  'markdown',
}

# Optional — enable the GFM task-list extension for rendered checkboxes:
myst_enable_extensions = ['tasklist']
```

Mixed `.rst` + `.md` projects work — see `examples/md_mixed/`.

### Word style template

To apply a custom Word style template, set `docx_template` in `conf.py`:

```python
# *.docx or *.dotx template file; resolved relative to the project srcdir
docx_template = 'template.docx'
```

### Debug logging

To capture visitor-level trace output for diagnosing rendering issues:

```python
docx_debug_log = True              # writes <outdir>/docx.log
# or
docx_debug_log = 'path/to/log'     # explicit path (relative to <outdir>)
```

Without this setting the writer is silent.

## Examples

The `examples/` directory contains ready-to-build Sphinx projects:

| RST | Markdown |
|---|---|
| `sample_1` — default example (upstream) | `md_basic` — headings / paragraphs / bold / italic / code |
| `sample_2` — `make docx` baseline | `md_lists` — bullets / numbered / nested / deflist |
| `sample_3` — custom `docx_template` | `md_tasklist` — GFM `- [ ]` / `- [x]` → `☐` / `☒` |
| `sample_4` — variation | `md_tables` — GFM pipe tables |
| `sample_5` — styled | `md_code` — fenced code blocks (python / bash / json) |
| `sample_6` — styled | `md_images` — image embedding |
| | `md_links` — external URLs and intra-doc anchors |
| | `md_admonitions` — MyST `:::{note}` / `:::{warning}` |
| | `md_mixed` — RST and MD source files in one project |

Build any example:

```
cd examples/md_tasklist
sphinx-build -b docx source build
# or, from the repo root:
make example DIR=md_tasklist
```

## Development

All development and testing happens in a Docker container. The host only
needs Docker and `make`:

```
make build    # build the docxsphinx-dev:py3.12 image
make test     # run the full pytest suite in the container
make lint     # ruff check src tests
make shell    # interactive shell in the container
```

Host-native escape hatch (bring your own virtualenv):

```
pip install -e ".[dev]"
make test LOCAL=1
```

Override the Python version:

```
PYTHON_VERSION=3.11 make build
PYTHON_VERSION=3.11 make test
```

### Test pyramid

The suite is split into four tiers, each selectable via `pytest -m <marker>`
or a Make target:

| Tier | Marker | Target | What it does |
|---|---|---|---|
| Smoke | `smoke` | `make test-smoke` | Import, builder registration, `setup()` metadata, entry point |
| Unit | `unit` | `make test-unit` | Hand-built doctrees → visitor → `Document`, no subprocess |
| Integration | `integration` | `make test-integration` | Parsed RST / MD → visitor → `Document`, in-memory inspection |
| End-to-end | `e2e` | `make test-e2e` | Full `sphinx-build` subprocess against each example, zip-validate output |

Inner-loop target: `make test-fast` (everything except e2e).

### Release

Tags matching `v*` trigger `.github/workflows/release.yml`, which builds
the sdist + wheel, runs `twine check`, and publishes to PyPI via
[Trusted Publishing](https://docs.pypi.org/trusted-publishers/) (OIDC —
no stored tokens).

Before the first release, configure a Trusted Publisher on PyPI pointing
at:

- Repository: `mherkazandjian/docxsphinx`
- Workflow: `release.yml`
- Environment: `pypi` (create this environment in the repo Settings → Environments)

### Profiling

```
make profile DIR=sample_2   # cProfile the writer against an example
```

Reports the top `writer.py` call sites by invocation count.

### Debug tip

```
python -m pdb $(which sphinx-build) -b docx path/to/src path/to/build
```

## License

MIT. See [`LICENSE`](LICENSE).
