---
name: docxsphinx-developer
description: Use when changing docxsphinx itself — the Sphinx docx builder (builder.py), the DocxTranslator visitor (writer.py), the raw-OOXML helpers (_docx_helpers.py), the version shims (_compat.py) or the docx2md / docx2rst reverse pipeline (reverse/) — or when adding tests, example projects or golden fingerprints, debugging a rendering bug, or preparing a release. Covers the architecture, the Docker-first dev loop, the five test tiers, the recipe for supporting a new docutils / Sphinx node, known gaps, and python-docx / OOXML gotchas. For configuring docxsphinx inside someone's Sphinx project, use docxsphinx-user instead.
---

# Developing docxsphinx

## Map

```
src/docxsphinx/
├── __init__.py       setup(): add_builder + config values docx_template / docx_debug_log /
│                     docx_documents; returns parallel_read_safe / parallel_write_safe
├── builder.py        DocxBuilder: one pass per output entry -> assemble doctree -> DocxWriter -> save
├── writer.py         DocxWriter (opens / resolves the template), DocxState, and DocxTranslator,
│                     the visitor (~1800 lines; nearly all feature work happens here)
├── _docx_helpers.py  raw-OOXML emitters python-docx lacks: hyperlinks, bookmarks, SEQ fields,
│                     the footnotes part, LaTeX -> OMML through pandoc
├── _compat.py        every Sphinx / docutils / python-docx import that has moved or might
└── reverse/          docx2md / docx2rst: engine.py (pandoc subprocess, typed errors),
                      cli.py (argparse only, no logic), __init__.py (public API)
tests/                five tiers, see below; golden fingerprints in tests/golden/
examples/             sample_* (RST), md_* (MyST), md_showcase = every feature in one file
tools/roundtrip.py    research harness: .docx -> pandoc -> MD -> docxsphinx -> .docx, diff OOXML
```

## How a build flows

```
sphinx-build -b docx
└─ DocxBuilder.write()
   └─ for entry in _document_entries():        # docx_documents, or one entry built from
      │                                         #   master_doc + "{project}-{version}"
      ├─ assemble_doctree(startdoc, toctree_only)
      │    inline_all_toctrees -> env.resolve_references -> fix_refuris
      ├─ DocxWriter(builder, template_override) # a fresh python-docx Document for each output
      ├─ writer.write(doctree)                  # DocxTranslator.walkabout
      └─ writer.save(<outdir>/<targetname>.docx)
```

Every source is folded into one doctree, so cross-document references end up as anchors within
the document (`fix_refuris` strips double anchors). `get_outdated_docs()` always returns
everything.

## The translator in one screen

- Node type `foo` dispatches to `visit_foo` / `depart_foo`. A visitor does one of three things:
  renders output; passes through (no output of its own, children still visited); or
  `raise nodes.SkipNode`, which **drops the node and its children**.
- A node type with no visitor hits `unknown_visit`. It logs one `docx.unknown_node` warning per
  node type per build (the set is stored on the builder), then **descends** so the text
  survives, unformatted. A block-level `TextElement` gets its own paragraph first. If
  descending a node is wrong (duplicate text, non-content), give it an explicit visitor
  rather than changing the default.
- Never raise SkipNode from a `depart_*`. Docutils only catches it around the visit, so once
  that node is walked a SkipNode from depart propagates out of `walkabout`.
- Two loggers: `logger` (stdlib `docxsphinx.writer`) is what `dprint` and `docx_debug_log`
  use, and its warnings bypass `-W` / `suppress_warnings`. `sphinx_logger`
  (`sphinx.util.logging`) is for warnings the user must act on: pass `location=node,
  type='docx', subtype=…`.
- Where output goes is a stack. `self.current_state` is a `DocxState` whose `.location` is the
  Document or a table `_Cell`. `new_state(cell)` pushes and `end_state()` pops
  (`self.old_states`). Tables, admonitions, definition / field / option lists push a cell state
  so that nested content lands inside the cell. The same object holds the table plumbing:
  `table`, `row`, `cell_counter`, `row_index`, `rowspan_until`, `pending_vmerges` (row spans are
  applied in `depart_table`).
- Runs go into `self.current_paragraph`. `add_text()` applies the inline flags `strong`,
  `emphasis`, `subscript`, `superscript` and `strike`, and drains `_pending_text_prefix`, which
  is how the tasklist glyph gets in.
- Lists: the `self.list_style` stack (`bullet` / `number`) plus `self.list_level` choose
  `List Bullet` / `List Number N`. Headings: `self.sectionlevel`. Numbering and anchors:
  `_bookmark_counter` / `_emitted_bookmarks`, `_figure_counter` / `_figure_numbers`,
  `_footnote_counter` / `_footnote_ooxml_id` / `_footnotes_root`.
- Links and footnote bodies keep their inline formatting by going through
  `_collect_styled_chunks` / `_walk_inline`. These produce `(text, bold, italic, strike)` tuples,
  which `add_hyperlink`, `add_internal_hyperlink` and `add_footnote` accept.
- Style lookups follow one pattern: `styles.get_style_id(name, WD_STYLE_TYPE.X)`, and on
  `KeyError` call `logger.warning` and use `None`. User templates may lack any style; never let
  that fail the build.
- Aliases are fine: `visit_number_reference = visit_reference`; every admonition, plus `topic`
  / `sidebar` / `todo_node`, maps to `_visit_admonition`, which turns a leading `title` child
  into the bold label and removes it so `visit_title` doesn't emit a Word heading; nodes with
  no look of their own (`container`, `legend`, `versionmodified`, `acks`, `desc_sig_*`) map to
  `_visit_passthrough`. Reuse `self.visit_paragraph(node)` to open a block paragraph with the
  usual list / table-cell reuse rules.

## Dev loop (Docker-first)

The host needs only Docker and `make`. The image is `python:<ver>-slim` with pandoc and
`pip install -e ".[dev]"`. The repo is bind-mounted at `/workspace` with
`PYTHONPATH=/workspace/src`, so source edits take effect without a rebuild. Rebuild the image
only after changing `pyproject.toml` or the Dockerfile.

```bash
make build                    # image (PYTHON_VERSION=3.11 make build for another Python)
make test-fast                # smoke + unit + integration: the inner loop
make test                     # all tiers, including e2e + golden
make test-smoke | test-unit | test-integration | test-e2e | test-golden
make golden-update            # rewrite tests/golden/*.fp.txt after an intended output change
make lint                     # ruff check src tests
make cov                      # coverage report
make example DIR=md_showcase  # or N=2 for sample_2; output in examples/<dir>/build/
make profile DIR=md_basic     # cProfile call counts for writer.py
make roundtrip [CORPUS=corpus]  # research harness; reports in examples/corpus/reports/ (gitignored)
make release-check            # python -m build + twine check
make shell                    # bash in the container
docker compose run --rm dev pytest -v tests/test_visitor_units.py::test_bold_run_is_bold
```

`LOCAL=1` runs any target in the host venv (`pip install -e ".[dev]"` first). Without pandoc on
the host, the reverse integration / e2e tests and the `md_math` golden **fail** rather than
skip.

## Tests: five tiers

| Tier | Where | Use for |
|---|---|---|
| `smoke` | `test_smoke.py` | Import, builder registration, entry point, `setup()` metadata |
| `unit` | `test_visitor_units.py` | A visitor on a hand-built doctree: `_blank_document()` (wraps `docutils.utils.new_document`; a bare `nodes.document(settings, None)` has no reporter), `_walk(rst, fake_builder)`, `_walk_doctree(doctree, fake_builder)` |
| `unit` | `test_helpers_units.py` | `_docx_helpers` functions in isolation, at the XML level |
| `unit` / `integration` / `e2e` | `test_reverse.py` | pandoc engine, with `subprocess.run` / `shutil.which` patched for unit; real pandoc otherwise |
| `integration` | `test_integration.py` | Real RST or MyST parsed through the `translator_factory(src, 'rst' or 'md')` fixture; catches parser-to-writer seam bugs |
| `e2e` | `test_build_examples.py` | A `sphinx-build` subprocess for every entry in `EXAMPLES`; exit code + zip validity |
| `e2e` | `test_dropped_content.py` | Every `…MARK` word in `sample_directives` reaches the docx; SVG / unknown-node warnings and `suppress_warnings` on a throwaway project |
| `golden` (+`e2e`) | `test_golden.py` | Rebuilds each example and diffs `tests/_fingerprint.py` output against `tests/golden/<example>.fp.txt` |

A fingerprint is a text summary of a `.docx`: counts, style distribution, headings, table shapes,
links, bookmarks and a tag census. `python tests/_fingerprint.py a.docx b.docx` prints one per
file; diff those outputs to compare two builds by hand.

## Recipe: support a construct (or stop dropping one)

1. **Look at the real doctree.** The two parsers produce different shapes for the same idea.
   ```bash
   docker compose run --rm -T dev python -c "
   from docutils.core import publish_doctree
   from myst_parser.parsers.docutils_ import Parser
   print(publish_doctree(open('/workspace/x.md').read(), parser=Parser(),
         settings_overrides={'myst_enable_extensions': ['colon_fence', 'deflist']}).pformat())"
   ```
   Leave out `parser=` for RST. Sphinx-only nodes (`desc_*`, `versionmodified`, `seealso`,
   `number_reference`, `todo_node`, …) exist only after Sphinx's transforms have run, so for
   those use `sphinx-build -b pseudoxml source build/pseudo` on an example.
2. **Find the visitor:** `grep -n "def visit_<tagname>" src/docxsphinx/writer.py`. Often there's
   a SkipNode stub, sometimes with commented-out code from the upstream text builder that's
   unrelated to Word. If there's no visitor at all, `unknown_visit` is rendering its bare text
   and the build log shows a `[docx.unknown_node]` warning for it.
3. **Implement** it next to its siblings. Start each method with `dprint()`, as every existing
   visitor does.
   - Block output: `self.current_paragraph = self.current_state.location.add_paragraph(style=…)`.
   - Container output: create a table / cell, `new_state(cell)` in visit, `end_state()` in
     depart.
   - Raw XML: add a helper to `_docx_helpers.py` (and its `__all__`) built with `OxmlElement` /
     `qn`. See [references/ooxml-and-python-docx.md](references/ooxml-and-python-docx.md).
   - When MyST and RST shapes differ, handle both in the writer; `_code_language()` is the model.
4. **Tests:** a unit test on a hand-built doctree, plus an integration test from RST and/or MyST
   source. A raw-XML helper also gets its own `test_helpers_units.py` test.
5. **Example:** extend the closest one (`md_showcase` for MyST features), or add
   `examples/<name>/{Makefile,source/conf.py,source/index.*}` and register it in `EXAMPLES` in
   `tests/test_build_examples.py`. `test_golden.py` imports the same list.
6. **Goldens:** `make golden-update`, then `git diff tests/golden/`. Every changed line should
   follow from your change; anything else is a regression.
7. `make lint && make test`.
8. **Docs:** the README "What it handles" table; `CHANGELOG.md` under `## [Unreleased]`
   (Keep a Changelog: Added / Fixed / Changed); and the user skill's
   `agents/skills/docxsphinx-user/references/feature-support.md` and its gap list, whenever
   behaviour that users can see changes.

## Debugging a rendering bug

- Reproduce on the smallest example: `make example DIR=<name>`.
- Read the XML. The image has no xmllint, so:
  ```bash
  docker compose run --rm -T dev python -c "import sys, zipfile, xml.dom.minidom as m; \
    print(m.parseString(zipfile.ZipFile(sys.argv[1]).read('word/document.xml')).toprettyxml())" \
    examples/<dir>/build/<file>.docx
  ```
  Footnotes are in `word/footnotes.xml`; relationships (hyperlinks, images) in
  `word/_rels/document.xml.rels`.
- Trace: `dprint()` with no arguments does nothing. Add `dprint(node=node.tagname, x=…)` in the
  visitor, set `docx_debug_log = True` in the example's `conf.py`, rebuild, then read
  `examples/<dir>/build/docx.log`. Revert both before committing. The handler is attached once
  per process.
- Step through it: `python -m pdb $(which sphinx-build) -b docx source build`.

## Known gaps (verified on main after 2.1.0, update this list as they close)

- `visit_literal` does nothing, so inline code isn't monospace.
- `visit_literal_block` sets `in_literal_block = True` and then raises SkipNode, so
  `depart_literal_block` never resets it. After the first code block, `visit_Text` keeps
  source newlines, and every wrapped paragraph gets hard `w:br` line breaks. Fixing it will
  shift most goldens.
- No visitor calls `dprint()` with arguments, so `docx_debug_log` produces an empty log, even
  though the README promises a trace of every visitor call.
- Titles outside sections, admonitions and topics (e.g. a table title) still go through
  `visit_title` and become Word headings.
- `desc_type_parameter_list` / `desc_type_parameter` (PEP 695 `[T]`) have no visitor: they
  warn and render as bare text.
- Setting `supported_image_types` on the builder would let `sphinx.ext.imgconverter` turn SVG
  into PNG, but it also switches on Sphinx's built-in `ImageDownloader` (remote images fetched
  at build time) and `DataURIExtractor`. That's not done yet.
- Image `:scale:` and `%` widths are ignored (they would need the intrinsic image size).
- Definition-list terms are flattened with `astext()`, which loses inline formatting.
  Classifiers are dropped.
- Not implemented: Word TOC field, page breaks, headers / footers, native checkbox content
  controls.

## Gotchas

- MyST puts the code-block language in `classes=['code', 'python']` and leaves
  `node['language']` as `None`. Use `_code_language(node)`.
- MyST strikethrough is `<raw><s></raw>text<raw></s></raw>` as **siblings**, with no wrapping
  node. `visit_raw` and `_collect_styled_chunks` toggle on the markers.
- MyST tasklists emit a raw-HTML `<input>`; `_checkbox_prefix` looks for `checked` in it.
- `literal_strong` / `literal_emphasis` live in `sphinx.addnodes`. Compare `tagname` strings
  rather than `isinstance`.
- `option_argument.astext()` already contains the delimiter; don't add `' '` or `'='` again.
- `.dotx` templates: python-docx refuses their content type. `DocxWriter._open_template`
  rewrites it in an in-memory copy of the zip.
- The template body is never cleared; output is appended after it.
- `add_table(style=…)` is only applied at the document root. Nested tables get no style.
- python-docx quirks (merged cells in `row.cells`, `run.text == '\n'`, `paragraph.text` skipping
  hyperlinks and fields, no footnote API, `parse_xml` refusing an XML declaration): see the
  [reference](references/ooxml-and-python-docx.md).

## Conventions

- ruff: line length 100, target py310, rules `I B UP SIM`. Each module starts with
  `from __future__ import annotations`.
- Import anything from upstream that has moved, or is at risk of moving, via `_compat.py`.
- The writer stays parser-agnostic. Dispatch on doctree shape, never on the source format.
- Degrade, don't crash: a cosmetic problem (missing style, math without pandoc, bad image) gets a
  `logger.warning` (from `sphinx.util.logging`) and a fallback.
- `reverse/engine.py` has no CLI or stdout concerns; `reverse/cli.py` holds no logic.

## Release

1. Bump `version` in `pyproject.toml`. Move `## [Unreleased]` in `CHANGELOG.md` to
   `## [X.Y.Z] — YYYY-MM-DD`.
2. Run `make release-check`.
3. Pushing a `vX.Y.Z` tag triggers `.github/workflows/release.yml`. `build` always runs (sdist,
   wheel, `twine check`). `publish` uses PyPI Trusted Publishing (OIDC, `pypi` environment) and
   runs only when the ref is a tag **and** the repo variable `PYPI_PUBLISH_ENABLED` is `'true'`
   (`gh variable set PYPI_PUBLISH_ENABLED --body true`).
4. For a dry run of the build half, use `gh workflow run release.yml --ref main`.

CI (`.github/workflows/ci.yml`, on pushes to `main` and on PRs) runs ruff and pytest with
coverage on Python 3.10–3.13, with pandoc installed. A separate `docker-smoke` job builds the
image with the runner's UID/GID, runs the suite in it, and builds `md_tasklist`.
