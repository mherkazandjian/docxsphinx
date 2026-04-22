# Changelog

All notable changes to this project are recorded in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.0] — _unreleased_

Breaking, consolidating release. Drops Python 2-era packaging, modernises
the dependency pins, formalises a four-tier test pyramid, adds Markdown
(MyST) input support alongside reStructuredText, and renders GFM task-list
checkboxes as `☐` / `☒` glyphs in the Word output.

### Added

- **PEP 621 packaging** via `pyproject.toml`: declarative metadata,
  dependency set, dev extras, ruff config, pytest config, and
  `sphinx.builders` entry point. `setup.py`/`setup.cfg`/`tox.ini`/
  `requirements.txt`/`MANIFEST.in` are removed.
- **`src/docxsphinx/_compat.py`** — single-file shim for every Sphinx,
  docutils, and python-docx import that has moved (or plausibly could)
  between versions. `builder.py` and `writer.py` import only through it.
- **`docx_debug_log` config value** — opt-in `FileHandler` attached to the
  `docxsphinx.writer` logger at DEBUG level. Default path is
  `<outdir>/docx.log` (not CWD). When unset, the writer is silent beyond
  Sphinx's own logging.
- **Dockerized dev environment**:
  - `Dockerfile` based on `python:${PYTHON_VERSION}-slim` (default 3.12)
    with a non-root `dev` user whose UID/GID match the host.
  - `docker-compose.yml` with a bind-mounted workspace.
  - `.dockerignore`.
  - `Makefile` driven by `docker compose run --rm dev` with a `LOCAL=1`
    escape hatch; targets `build`, `shell`, `test`, `lint`, `cov`,
    `example N=<num> | DIR=<dir>`, `profile`, `release-check`, `clean`,
    `distclean`, plus tier-scoped `test-smoke` / `test-unit` /
    `test-integration` / `test-e2e` / `test-fast` / `test-all`.
- **GitHub Actions CI** (`.github/workflows/ci.yml`):
  - `lint-test` matrix — Python 3.10 / 3.11 / 3.12 / 3.13, `ruff check`
    + `pytest --cov`, coverage artifact from the 3.12 job.
  - `docker-smoke` — `docker compose build` + full pytest suite + a
    `make example` sanity check inside the container.
- **GitHub Actions release workflow** (`.github/workflows/release.yml`):
  on tag `v*`, build sdist + wheel, run `twine check`, publish to PyPI via
  Trusted Publishing (OIDC — no stored tokens). Publish gated on a `pypi`
  GitHub environment.
- **Four-tier test pyramid** with markers registered in `pyproject.toml`:
  - `tests/test_smoke.py` — import, builder registration, `setup()`
    metadata, `sphinx.builders` entry-point resolution.
  - `tests/test_visitor_units.py` — hand-built doctree unit tests.
  - `tests/test_integration.py` — visitor + python-docx on real RST and
    MD source, inspected in-memory.
  - `tests/test_build_examples.py` — end-to-end `sphinx-build`
    subprocess; validates the emitted `.docx` is a zip containing a
    well-formed `word/document.xml`.
  - `tests/conftest.py` — shared fixtures (`repo_root`,
    `examples_root`, `fake_builder`, `translator_factory`);
    `sphinx.testing.fixtures` registered as a plugin.
- **Markdown (MyST) input support**. `myst-parser>=3.0` in dev extras;
  the writer is parser-agnostic, so enabling MyST in a user's `conf.py`
  works with no additional docxsphinx configuration.
- **GFM task-list / checkbox rendering**. A list_item carrying the
  `task-list-item` class emits `☐` (unchecked) or `☒` (checked) as a
  leading glyph in the same `List Bullet` paragraph as the task text.
- **Admonition rendering**. `note`, `warning`, `tip`, `caution`,
  `important`, `hint`, `error`, `danger`, `attention`, and generic
  `.. admonition:: Custom heading` directives each render as a 1×1
  table whose cell opens with a bold title paragraph and contains the
  admonition body. Previously every admonition type raised `SkipNode`
  and the content was silently dropped.
- **Subscript / superscript rendering**. `subscript` and `superscript`
  nodes (from RST `:sub:` / `:sup:` roles or equivalent MyST markup)
  now produce runs with `font.subscript` / `font.superscript` set,
  rather than being dropped via `SkipNode`.
- **Image sizing and alignment**. `image` nodes honour `:width:`,
  `:height:` (px, cm, mm, in, pt), `:align:` (center/left/right), and
  `:alt:` attributes. Inline images within paragraphs go into the
  current paragraph as a run; block images get their own paragraph
  whose alignment reflects `:align:`. Percentage widths and `:scale:`
  emit a warning and fall back to native sizing. Alt text propagates
  to the inline shape's `wp:docPr` `descr`/`title` attributes.
- **Definition / glossary list rendering**. `definition_list` is
  emitted as a two-column table: term (bold) in column 0, definition
  content (with full inline/block formatting, including multiple
  paragraphs per definition) in column 1. MyST's multi-`:` pattern
  (`Term\\n: def A\\n: def B`) lands as stacked paragraphs within the
  single term's definition cell. Previously the whole directive
  raised `SkipNode`.
- **Option list rendering**. `option_list` emits a two-column table:
  the canonical option signature (`-o FILE, --out=FILE` — flattened
  from the `option_group`/`option_string`/`option_argument` subtree)
  in column 0, the description body in column 1. Multi-paragraph
  descriptions render correctly.
- **Table rowspan (`morerows`)**. `visit_entry` previously raised
  `NotImplementedError` when a cell declared `morerows=N`. It now
  records a vertical-merge plan, skips column slots in subsequent
  rows that are still occupied by earlier rowspans (so `cell_counter`
  advances correctly when the doctree omits `<entry>` children for
  spanned columns), and emits the merge in `depart_table` once all
  rows exist. Combined `morerows + morecols` (rectangular merges)
  are supported.
- **External hyperlinks**. `reference` nodes carrying a `refuri`
  with a known scheme (`http`/`https`/`ftp`/`mailto`/`file`/`data`/
  `tel`) render as clickable `w:hyperlink` elements. The anchor text
  is blue and underlined regardless of the template (styling is
  applied directly on the run, not via a named style), and the URL
  is registered as an external relationship on the paragraph's part.
  Implementation lives in a new `src/docxsphinx/_docx_helpers.py`
  module.
- **Internal hyperlinks and section bookmarks**. Every `section`
  node's `ids` now emit `w:bookmarkStart`/`End` zero-width bookmarks
  in the heading paragraph. `reference` nodes carrying a `refid` (or
  a fragment-only `refuri` like `#my-section`) render as
  `w:hyperlink` elements with `w:anchor=<id>` — clickable jumps to
  the matching bookmark. The `md_links` example now emits 8 section
  bookmarks and at least one working internal anchor link out of the
  box.
- **Code-block syntax highlighting (Pygments)**. `literal_block`
  emits one coloured run per Pygments token, with a small palette for
  keyword / name / string / number / comment / generic-* categories
  inlined on each run (so no template Hyperlink/Code style needed).
  Newlines within the source become `w:br` elements so the block
  wraps correctly inside a single paragraph. Language detection
  reads `node['language']` (RST `code-block::`) or — MyST's
  convention — the first non-`code` entry of `node['classes']`.
  Unknown/missing languages fall back to Pygments' `TextLexer`
  (plain monospace). All runs get the `Consolas` face.
- **Figure captions with auto-numbering**. `visit_caption` now emits
  a dedicated `Caption`-styled paragraph prefixed by a bold `Figure
  <N>. ` where `<N>` is produced by a Word `SEQ Figure \* ARABIC`
  field (auto-renumbers on field refresh; a 1-based placeholder is
  filled in immediately for viewers that don't update fields).
  Figure ids get matching zero-width bookmarks in the caption, so
  cross-references can anchor at the caption paragraph. A
  translator-level `figure_id → number` map is maintained for
  future `:numref:` support.
- **Footnotes and citations**. `footnote` and `citation` bodies write
  into a newly-created `word/footnotes.xml` part (created on demand
  via `ensure_footnotes_part`; the OPC relationship is registered
  on the main document). Inline `footnote_reference` and
  `citation_reference` nodes emit superscripted
  `<w:footnoteReference w:id="N"/>` runs whose ids match the
  corresponding `<w:footnote>` body entries. OOXML ids are allocated
  lazily and keyed by the target footnote's docutils id so references
  and bodies always pair up regardless of visit order.
- **Strikethrough (MyST `~~text~~`)**. The writer now toggles a
  strike-run flag on raw-HTML `<s>` / `<strike>` / `<del>` tag
  markers in the doctree (MyST emits those as sibling `<raw>` nodes
  around the struck text after warning that strikethrough only
  renders natively in HTML). Struck text emits with
  `run.font.strike = True`.
- **Inline formatting inside hyperlinks and footnote bodies**. The
  `add_hyperlink` / `add_internal_hyperlink` / `add_footnote`
  helpers now accept either a plain string (single run) *or* a list
  of `(text, bold, italic, strike)` chunks. The writer walks the
  source reference/footnote subtree via `_collect_styled_chunks` —
  a dedicated inline walker that recognises `strong`,
  `literal_strong`, `emphasis`, `literal_emphasis`, and the MyST
  strikethrough raw-HTML sibling markers — and passes the chunks
  through. `[**bold link**](url)` now renders with an actually-bold
  hyperlink run; footnote bodies carrying `**bold**` / `*italic*` /
  `~~strike~~` round-trip correctly.
- **Nine new example projects** under `examples/md_*/`, each a minimal
  Sphinx project with a matching `source/conf.py` and `source/index.md`:
  `md_basic`, `md_lists`, `md_tasklist`, `md_tables`, `md_code`,
  `md_images`, `md_links`, `md_admonitions`, `md_mixed` (RST + MD in
  one project).
- `CHANGELOG.md` (this file).

### Changed

- **Python minimum raised to 3.10.** Previously no minor was pinned in
  classifiers, Pipfile declared 3.6, tox listed `py27,py36`.
- **Dependency pins modernised**:
  - `docutils`: `==0.15` → `>=0.17`
  - `Sphinx`: `>=3.0.0` → `>=3.5`
  - `python-docx`: `>=0.8.0` → `>=1.0.0`
  - `Pygments>=2.14` added as a direct dep (reserved for the upcoming
    code-highlighting feature).
- **Writer silent by default.** The module-level
  `logging.basicConfig(filename='docx.log', ...)` call is gone; the
  writer no longer writes a `docx.log` to CWD unless the user
  opts in via `docx_debug_log` in `conf.py`. `dprint()` became a no-op
  when called without kwargs (the common case) and emits at DEBUG
  otherwise.
- **Template path resolution fixed.** `docx_template = 'template.docx'`
  is now resolved against `builder.env.srcdir`. Previously templates
  only loaded when `sphinx-build` was invoked from the sample project's
  parent directory (`os.path.join('source', dotx)` relative to CWD).
- `DocxWriter.template_dir = "NO"` sentinel replaced with
  `self.template_path: Path | None` set in `__init__`.
- `DocxState(object):` → `DocxState:` (Python 3 idiom).
- `DocxTranslator.__init__`: added `_pending_text_prefix`, drained by
  `add_text` on the next text emission — the mechanism behind task-list
  glyph rendering.
- `tests/test_build_examples.py` rewritten: uses `subprocess.run(...)`
  and asserts `returncode == 0` with stderr surfaced in the failure
  message. The legacy version discarded the return code via
  `Popen(...).wait()`, so sphinx-build failures that still left an
  output file silently passed.
- `make example` accepts either `N=<num>` (legacy `sample_N` paths) or
  `DIR=<name>` (the new Markdown example paths).
- Imports modernised: `Callable` moved from `typing` to
  `collections.abc`; import blocks sorted by ruff's `I` rules.

### Fixed

- `visit_paragraph` crashed with `AttributeError: 'NoneType' object has
  no attribute 'style'` when a doctree started with a bare paragraph
  (no section/title first). Added a None-guard; surfaced by the new
  unit-test tier.
- Numbered (`enumerated_list`) items incorrectly rendered with the
  `List Bullet` style. The pre-existing but unused `self.list_style`
  stack is now maintained in `visit_bullet_list` / `visit_enumerated_list`
  and consulted in `visit_list_item`, so numbered items render with
  `List Number` (or `List Number 2`/`3`/… at deeper levels). Mixed
  nesting picks the innermost list's kind correctly.
- `E741` ambiguous single-letter `l` variable in
  `visit_tabular_col_spec`; renamed to `piece`.
- `F841` unused `col` local in `visit_colspec`.
- Sample-project `conf.py` files (samples 2/3/4) were incompatible with
  Sphinx >= 7:
  - `language = None` → `language = 'en'`
  - `intersphinx_mapping = {'https://docs.python.org/': None}` →
    `{'python': ('https://docs.python.org/3', None)}`

### Removed

- `setup.py`, `setup.cfg`, `tox.ini`, `requirements.txt`, `MANIFEST.in`
  (superseded by `pyproject.toml`).
- `.circleci/config.yml` (superseded by GitHub Actions).
- The unconditional module-level `logging.basicConfig` call in
  `writer.py` that wrote `docx.log` to CWD.
- Stale `docx.log` files inside each `examples/sample_*/` directory
  (leftovers from the removed module-level logging).
- `src/TODO.txt` (stale backlog — tracking moved to issues / the
  CHANGELOG's unreleased section).
- `src/README.md` (outdated — held the upstream `sphinxcontrib-docxbuilder`
  description, superseded by the repo-root `README.md`).
- `src/LICENSE.txt` — promoted verbatim to a repo-root `LICENSE` file so
  setuptools' automatic license-metadata mechanism picks it up (it now
  ships as `dist-info/licenses/LICENSE` in the wheel and `LICENSE` at
  the sdist root). The fork-author copyright line is appended to
  the upstream one.
- `Pipfile` and `Pipfile.lock` — the old Pipfile pinned `docutils==0.15`
  / `Sphinx>=3.0.0` and was replaced by `pyproject.toml`; the lock
  file had stayed 19 KB of stale pins for two years. Docker-first
  dev (`make build && make test`) and bare-metal `pip install -e
  ".[dev]"` both work without pipenv.

## [1.0.1] — 2022

- Version bump after 1.0.0 release. No functional changes.

## [1.0.0] — 2022

- First tagged release of the `mherkazandjian/docxsphinx` fork. Heavy
  modification of the upstream
  [sphinxcontrib-docxbuilder](https://bitbucket.org/shimizukawa/sphinxcontrib-docxbuilder)
  by Mher Kazandjian and Hugo Buddelmeijer. See the README for earlier
  history.
