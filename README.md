# docxsphinx

[![CI](https://github.com/mherkazandjian/docxsphinx/actions/workflows/ci.yml/badge.svg)](https://github.com/mherkazandjian/docxsphinx/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/docxsphinx.svg)](https://pypi.org/project/docxsphinx/)
[![Python](https://img.shields.io/pypi/pyversions/docxsphinx.svg)](https://pypi.org/project/docxsphinx/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

**Sphinx ↔ Word, both directions.**
`docxsphinx` is a Sphinx builder extension that emits Microsoft Word
(`.docx`) from Sphinx documentation sources, and — as of 2.1 — ships a
companion reverse pipeline (`docx2md`, `docx2rst`) that converts Word
documents back to Markdown or reStructuredText so teams with legacy
`.docx` content can adopt a Sphinx / Git workflow without rewriting from
scratch.

Both reStructuredText **and** Markdown (via
[myst-parser](https://myst-parser.readthedocs.io/)) are first-class input
formats. Mixed `.rst` + `.md` projects work out of the box.

Originally forked from
[shimizukawa/sphinxcontrib-docxbuilder](https://bitbucket.org/shimizukawa/sphinxcontrib-docxbuilder)
with substantial modification; see [`CHANGELOG.md`](CHANGELOG.md) for
release notes.

## What it handles

| Input construct | Output |
|---|---|
| Headings H1–H6 | Word heading styles, with section-id bookmarks for cross-refs |
| Paragraphs, bold, italic, inline code, sub/superscript, strikethrough | Run-level formatting on the matching character properties |
| Bullet lists (incl. nested) | `List Bullet` / `List Bullet 2` / … styles |
| Numbered lists (incl. nested, mixed-kind) | `List Number` / `List Number 2` / … styles |
| GFM task lists (`- [ ] …` / `- [x] …`) | `☐` / `☒` glyphs inline in the list paragraph |
| Definition / glossary lists | 2-column table, bold term + definition body |
| Option lists (`-h, --help  description`) | 2-column table with canonical flag signature + description |
| Tables, incl. column + row spans | Native Word tables with `w:gridSpan` + `w:vMerge` |
| Images with `:width:` / `:height:` / `:align:` / `:alt:` | Inline shapes sized correctly in EMU, alt text propagated |
| Figures with captions | `Caption`-styled paragraph with auto-numbered `SEQ Figure` field |
| Admonitions — `note` / `warning` / `tip` / `caution` / `error` / `danger` / `important` / `hint` / `attention` + generic | 1×1 wrapping table with bold title run + body |
| External hyperlinks | `w:hyperlink` with `r:id` + inline formatting inside link text |
| Internal references (`:ref:`, `[text](#id)`, `(label)=`) | `w:hyperlink` with `w:anchor` targeting auto-emitted section bookmarks |
| Numbered references (`:numref:` with `numfig = True`) | Clickable `w:hyperlink` carrying Sphinx's formatted label ("Fig. 1" / "Table 3" / etc.) |
| Fenced code blocks (python, bash, JSON, …) | Pygments-coloured monospace runs, line breaks via `w:br` |
| Footnotes and citations with inline formatting | Word-native `word/footnotes.xml` part + superscripted references |
| LaTeX math (`:math:`, `.. math::`, MyST `$…$` / `$$…$$`, AMS environments) | Native Office Math Markup Language (OMML) via pandoc; editable in Word's equation editor |
| `sphinx.ext.autodoc` (`.. automodule::`, `.. autoclass::`, `.. autofunction::`) | Function / method / class signatures with bold names, parameter lists, optional-arg `[…]` brackets, return-type arrows; docstring `:param:` / `:returns:` / `:raises:` field lists as 2-column tables |

Features not yet implemented (tracked on the roadmap):
versionadded/deprecated directives, Word-native TOC field, native
form-field checkboxes, page headers/footers/breaks.

## Installation

```
pip install docxsphinx
```

From source (default branch is `main`):

```
pip install git+https://github.com/mherkazandjian/docxsphinx.git@main
```

## Forward — Sphinx project → `.docx`

In your Sphinx project's `conf.py`:

```python
extensions = ['docxsphinx']
```

Then invoke the builder:

```
sphinx-build -b docx source build
```

The builder inlines every toctree into a single doctree and emits one
`.docx` named `{project}-{version}.docx` (from `conf.py`) inside the
build directory.

### Markdown input

Register MyST as the Markdown parser:

```python
extensions = ['myst_parser', 'docxsphinx']
source_suffix = {
    '.rst': 'restructuredtext',
    '.md':  'markdown',
}

# Enable whichever MyST extensions your docs use:
myst_enable_extensions = [
    'tasklist',        # GFM - [ ] / - [x]
    'colon_fence',     # :::{note} admonitions
    'deflist',         # term/definition lists
    'strikethrough',   # ~~struck~~
    'attrs_block',     # block-level attributes on images/figures
    'html_image',      # recognise <img> tags from pandoc output
]
```

Mixed `.rst` + `.md` in the same project works — see
[`examples/md_mixed/`](examples/md_mixed/).

### Configuration reference

| `conf.py` value | Default | Purpose |
|---|---|---|
| `docx_template` | `None` | Path to a `.docx` or `.dotx` template whose styles carry over. `.dotx` is loaded by transparently rewriting the main-document content-type override in memory (python-docx itself won't open `.dotx` — see #41). Relative paths are searched under each entry of `templates_path` (in order), then under `srcdir` itself. Absolute paths are used as-is. |
| `docx_debug_log` | `None` | `True` or a relative path → visitor-level DEBUG trace written to `<outdir>/docx.log` (or the given path). Default: writer is silent. |

### Example with a custom template

```python
# conf.py
extensions = ['docxsphinx']
docx_template = 'template.docx'    # sibling to conf.py
```

Place `template.docx` in `source/` alongside `conf.py` and its styles
(`Heading 1..4`, `Normal`, `List Bullet`, `List Bullet 2`, `List Number`,
`Grid Table 4`, `Preformatted Text`, `Caption`, `FootnoteReference`)
become the palette docxsphinx renders into. Missing styles fall back to
no style rather than erroring.

## Reverse — `.docx` → Markdown / RST

Installed alongside the forward path, wraps [pandoc](https://pandoc.org/)
(2.0+) as a subprocess.

Install pandoc (once):

```
apt install pandoc                              # Debian / Ubuntu
brew install pandoc                             # macOS
winget install --id JohnMacFarlane.Pandoc       # Windows
```

Then:

```
pip install "docxsphinx[reverse]"

docx2md  report.docx -o report.md                        # default: GFM
docx2md  report.docx --format commonmark                 # or commonmark / markdown / markdown_strict
docx2rst report.docx -o report.rst

docx2md  --help                                          # full CLI options
```

By default, embedded images are extracted to a sibling directory named
`<output-stem>_media/` so the resulting Markdown rebuilds cleanly via
docxsphinx's forward path. Override with `--extract-media <dir>` or pass
`--extract-media -` to drop images entirely.

Programmatic use:

```python
from docxsphinx.reverse import docx_to_markdown, docx_to_rst

md = docx_to_markdown('report.docx', extract_media='media/')
rst = docx_to_rst('report.docx')
```

### Known limitations (MVP)

- Custom Word paragraph / character styles flatten to plain paragraphs.
- Pygments-coloured code blocks emitted by docxsphinx's forward path
  don't round-trip back to fenced code blocks — pandoc sees per-token
  runs as styled prose. Code-block *content* survives, fencing doesn't.
- Equations fall back to raw LaTeX.
- Floating images, tracked changes, comments, and form fields are
  dropped.
- Documents authored with zero explicit heading styles (all "Normal" +
  visual formatting) lose heading structure.

See [`docs/roundtrip-findings.md`](docs/roundtrip-findings.md) for the
per-feature round-trip diagnostic.

### Round-trip research harness

The companion research tool [`tools/roundtrip.py`](tools/roundtrip.py)
sends a corpus of real Word-authored documents through
`pandoc → Markdown → docxsphinx → .docx` and diffs the resulting OOXML
against the originals. The output drives the backlog of
forward-direction idiom improvements.

```
make roundtrip                        # against committed fixtures
make roundtrip CORPUS=corpus          # against your private corpus dropped into examples/corpus/
```

Reports land in `examples/corpus/reports/` (gitignored).

## Examples

Seventeen ready-to-build Sphinx projects live in [`examples/`](examples/):

| RST | Markdown |
|---|---|
| `sample_1` — upstream baseline | `md_basic` — headings / paragraphs / bold / italic / code |
| `sample_2` — plain `make docx` | `md_lists` — bullets / numbered / nested / deflist |
| `sample_3` — custom `docx_template` | `md_tasklist` — GFM `- [ ]` / `- [x]` → `☐` / `☒` |
| `sample_4` — variation | `md_tables` — GFM pipe tables |
| `sample_5` — styled template | `md_code` — fenced code blocks (python / bash / json) |
| `sample_6` — styled template | `md_images` — native + sized + aligned images |
| `sample_autodoc` — `sphinx.ext.autodoc` against an in-repo Python module | `md_links` — external URLs and intra-doc anchors |
| `sample_nested_toctree` — three-file toctree chain → H1/H2/H3 | |
| | `md_admonitions` — MyST `:::{note}` / `:::{warning}` / all types |
| | `md_mixed` — RST and MD source files in one project |
| | `md_math` — LaTeX equations → OMML (inline, display, AMS environments) |
| | **[`md_showcase`](examples/md_showcase/) — every feature in one file** |

Build any example:

```
make example DIR=md_showcase              # in the docker dev container
# or directly:
cd examples/md_showcase
sphinx-build -b docx source build
```

## Development

All development and testing runs inside a Docker container — the host
only needs Docker and `make`:

```
make build              # docker compose build (python:3.12-slim by default)
make test               # full pytest suite in the container
make lint               # ruff check src tests
make shell              # interactive bash in the container
make example DIR=<dir>  # build one example and leave the .docx in examples/<dir>/build/
make roundtrip          # run the reverse+forward research harness
make release-check      # python -m build && twine check
```

Host-native escape hatch (bring your own virtualenv):

```
pip install -e ".[dev,reverse]"
make test LOCAL=1
```

Override the Python version:

```
PYTHON_VERSION=3.11 make build
PYTHON_VERSION=3.11 make test
```

### Test pyramid

Four tiers, each selectable via a marker or a Make target:

| Tier | Marker | Make target | Focus |
|---|---|---|---|
| Smoke | `smoke` | `make test-smoke` | Imports, builder registration, `setup()` metadata, entry-point resolution |
| Unit | `unit` | `make test-unit` | Hand-built doctrees → visitor → `Document`, no subprocess |
| Integration | `integration` | `make test-integration` | Parsed RST / MD → visitor → `Document`, in-memory inspection |
| End-to-end | `e2e` | `make test-e2e` | `sphinx-build` subprocess per example, zip-valid output, `docx2md` round-trip |
| Golden | `golden` | `make test-golden` | Fingerprint regression against `tests/golden/<sample>.fp.txt` — catches silent content loss |

Inner-loop target: `make test-fast` (everything except e2e).

When a feature change intentionally alters the content fingerprint,
regenerate the goldens with `make golden-update` (sets `UPDATE_GOLDEN=1`
under the hood) and review the text diff in your PR.

### Profiling

```
make profile DIR=md_showcase          # cProfile the writer against an example
```

Reports the top `writer.py` call sites by invocation count.

### Debug tip

Drop `docx_debug_log = True` into the project's `conf.py` to capture a
DEBUG-level trace of every visitor call to `<outdir>/docx.log`. For
interactive debugging:

```
python -m pdb $(which sphinx-build) -b docx path/to/src path/to/build
```

### Release

Tags matching `v*` trigger [`.github/workflows/release.yml`](.github/workflows/release.yml),
which builds the sdist + wheel, runs `twine check`, and publishes to
PyPI via [Trusted Publishing](https://docs.pypi.org/trusted-publishers/)
(OIDC — no stored tokens).

Dry-run the workflow (build + `twine check`, no publish) without
tagging:

```
gh workflow run release.yml --ref main
gh run watch
```

The `publish` job is gated on a tag ref, so a manual trigger exercises
only the build half of the pipeline.

Before the first real release, configure a Trusted Publisher on PyPI:

- Repository: `mherkazandjian/docxsphinx`
- Workflow: `release.yml`
- Environment: `pypi` (create in Settings → Environments)

## License

MIT. See [`LICENSE`](LICENSE).
