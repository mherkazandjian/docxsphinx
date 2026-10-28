---
name: docxsphinx-user
description: Use when a Sphinx project should produce Microsoft Word (.docx) output with the docxsphinx extension, or when converting existing .docx files to Markdown or reStructuredText with docx2md / docx2rst. Covers installing, conf.py setup for reStructuredText and Markdown (MyST), docx_template / docx_documents / docx_debug_log, Word style templates and per-table styles, equations (need pandoc), which directives render and which are silently dropped, and troubleshooting missing styles, missing content, images and build errors. Not for changing docxsphinx's own source — use docxsphinx-developer for that.
---

# Using docxsphinx

docxsphinx is a Sphinx extension that registers a builder named `docx`. The builder folds every
toctree into one doctree and writes **one `.docx` per configured output**, by default a single
file named `{project}-{version}.docx`. It works on the docutils doctree, so reStructuredText,
Markdown (through MyST) and projects that mix both all take the same path. The package also
ships `docx2md` / `docx2rst` for the opposite direction.

Requirements: Python >= 3.10, Sphinx >= 3.5, python-docx >= 1.0. A `pandoc` binary (>= 2.0) on
`PATH` is optional: it is needed only for equations and for `docx2md` / `docx2rst`.

## Set up a project

```bash
pip install docxsphinx              # plus myst-parser if any sources are Markdown
```

reStructuredText-only `conf.py`:

```python
project = 'Handbook'
version = '1.2'                     # unset version gives "Handbook-.docx"
extensions = ['docxsphinx']
```

Markdown or mixed `conf.py`:

```python
extensions = ['myst_parser', 'docxsphinx']
source_suffix = {'.rst': 'restructuredtext', '.md': 'markdown'}
myst_enable_extensions = [
    'colon_fence',    # :::{note} admonitions
    'deflist',        # term / definition lists
    'tasklist',       # - [ ] / - [x]  ->  ☐ / ☒
    'strikethrough',  # ~~text~~
    'dollarmath',     # $inline$ and $$display$$ math
    'amsmath',        # \begin{align} ... \end{align}
    'attrs_block',    # {width=...} attributes on images
    'html_image',     # <img> tags, e.g. from docx2md output
]
```

Build:

```bash
sphinx-build -b docx source build   # -> build/Handbook-1.2.docx
sphinx-build -M docx source build   # -> build/docx/Handbook-1.2.docx  (this is what `make docx` runs)
```

`make docx` works in any project created by `sphinx-quickstart`, because its catch-all Makefile
target forwards to `sphinx-build -M`.

## Configuration values

| `conf.py` value | Default | Meaning |
|---|---|---|
| `docx_template` | `None` | `.docx` or `.dotx` whose styles (and body text, see below) the output starts from. |
| `docx_documents` | `[]` | One tuple per output file. Empty means one file from `master_doc`. |
| `docx_debug_log` | `None` | `True` attaches a DEBUG log file at `<outdir>/docx.log`. A path string writes there instead (relative paths resolve against the output dir). As of 2.1 the writer emits almost nothing to it; it's a hook for developers. |

### `docx_template`

- A relative path is searched for in each `templates_path` entry (relative to the source dir),
  then in the source dir itself. An absolute path is used as given.
- If it's not found, Sphinx logs `docx_template '…' not found; searched …`. The build then fails
  when python-docx tries to open the file. Fix the path, don't ignore the warning.
- `.dotx` works directly; there's no need to re-save it as `.docx`.
- **The template's body is kept.** Generated content is appended after whatever text the
  template already contains. That makes a cover page easy; leave the body empty if you don't
  want one.

### `docx_documents`: several outputs from one project

Same idea as Sphinx's `latex_documents`:

```python
docx_documents = [
    # (startdoc, targetname, template, toctree_only)
    ('guide/index', 'UserGuide', None, False),          # ".docx" is appended if missing
    ('admin/index', 'AdminGuide.docx', 'admin.dotx', False),  # per-output template
    ('index', 'Everything', None, True),
]
```

- `startdoc` is a docname (no suffix). It is the root of that output's toctree traversal.
- `template` overrides `docx_template` for this output only. `None` falls back to it.
- `toctree_only=True` keeps the start document's title and the chapters its toctrees pull in,
  and drops the start document's own prose. Use it when the master is just a cover or index.

## Templates and styles

The writer asks for these styles by name. A missing style logs
`style "X" is missing (…); falling back to no style`, and the content still renders.

| Content | Style | In python-docx's built-in default? |
|---|---|---|
| Section titles | `Heading 1`, `Heading 2`, … (by nesting depth, across toctrees) | yes |
| Bullet lists | `List Bullet`, `List Bullet 2`, `List Bullet 3`, … | yes |
| Numbered lists | `List Number`, `List Number 2`, … | yes |
| Tables | `Grid Table 4` | **no** |
| Code blocks | `Preformatted Text` (runs are Consolas and Pygments-coloured either way) | **no** |
| Figure captions | `Caption` | yes |

So without a template, tables come out with no borders and code blocks have no paragraph style.
To build a template, open Word, apply each style above at least once, and save. Word only writes
a built-in style into the file after it has been used. You can also set the style one table at a
time. A comment just before a table sets the style of the **next table only**:

```rst
.. DocxTableStyle Table Grid
```

```md
% DocxTableStyle Table Grid
```

The name has to exist in the template; python-docx's default includes `Table Grid` and
`Light Grid Accent 1`. Only real tables take a table style (grid, simple, `list-table`,
`csv-table`, GFM pipe tables). docxsphinx also builds definition lists, field lists, option lists
and admonitions out of tables, and those always use the template's default table style.

## What renders, and what disappears

The full matrix is in [references/feature-support.md](references/feature-support.md). Read it
before promising that a directive will show up in Word. These are the gaps people hit most:

- **Silently dropped, including their content, with no warning:** `.. contents::`,
  `.. topic::`, `.. sidebar::`, `.. container::`, `.. rubric::`, line blocks (`| …`), bare
  doctest blocks (`>>> …` without `::`), `versionadded` / `versionchanged` / `deprecated`,
  `.. todo::`, figure legends, block-quote attributions. The same happens to **any node type
  from another extension** that docxsphinx has no handler for.
- **Inline code** (``` ``x`` ``` / `` `x` ``) comes out as plain text, not monospace.
- **Images:** an SVG, or any format other than PNG / JPEG / GIF / BMP / TIFF, **aborts the
  build**, so convert it first. Remote `http(s)` images are skipped with a warning. `:scale:`
  and `%` widths are ignored with a warning; use px, cm, mm, in or pt.
- **Equations** need `pandoc` on `PATH` when you build. Without it, the raw LaTeX is printed in
  Consolas and a `math rendered as plain LaTeX` warning is logged.
- **Figure numbers** are Word `SEQ` fields with a placeholder value. In Word, press Ctrl+A then
  F9 to refresh them.
- There's no Word table of contents, page break, header or footer. Add them in Word after the
  build, or put them in the template. Headings use the real `Heading N` styles, so Word's
  References → Table of Contents works on the output.

## Reverse: `.docx` → Markdown / RST

These commands wrap pandoc, which must be installed on its own. The `[reverse]` extra adds no
Python packages.

```bash
docx2md  report.docx -o docs/report.md             # GFM; images go to docs/report_media/
docx2md  report.docx -o out.md --format commonmark  # or markdown / markdown_strict
docx2rst report.docx -o docs/report.rst
docx2md  report.docx --extract-media img/ > report.md   # stdout drops images unless --extract-media
```

Exit codes: `0` ok, `127` pandoc not on `PATH`, `2` pandoc older than 2.0 or input missing,
anything else is pandoc's own exit code.

```python
from docxsphinx.reverse import docx_to_markdown, docx_to_rst
md = docx_to_markdown('report.docx', flavour='gfm', extract_media='media/')
```

To migrate legacy Word content into Sphinx: run `docx2md` into `source/`, add the file to a
toctree, enable MyST with `html_image` (pandoc writes sized images as `<img>` tags), and build
with `-b docx`. What doesn't round-trip:

- custom paragraph and character styles become plain paragraphs;
- docxsphinx's own coloured code blocks come back as prose;
- equations come back as raw LaTeX;
- floating images, tracked changes, comments and form fields are lost;
- documents that never used heading styles lose their structure.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Builder name docx not registered or available through entry point` | docxsphinx is not installed in the Python that runs `sphinx-build`. Install it there and list it in `extensions`. |
| `style "Grid Table 4" is missing` / `"Preformatted Text" is missing` | No template, or the template never used that style. Add a template, or use `DocxTableStyle`. |
| Text from a directive is missing in Word | It's on the dropped list above. To see which node holds it, run `sphinx-build -b pseudoxml source build/pseudo`, open the page's `.pseudoxml` file and look up the enclosing node (`<container>`, `<line_block>`, `<versionmodified>`, …) in the feature matrix. Rewrite with a supported construct, e.g. `.. note::` instead of `versionadded`, a heading instead of `rubric`, `.. code-block:: pycon` instead of a bare `>>>` block. |
| Build aborts while writing, on an image | The image is SVG or another unsupported format. Convert it to PNG. |
| Image missing, with `failed to embed image` | Wrong path, or a remote URL. Image paths are relative to the document, and the file must be local. |
| Equations show up as LaTeX source | pandoc is missing on the build machine (in CI too). |
| `docx_template … not found; searched …` | The path is relative to `templates_path` entries or the source dir, not to where you run the build. |
| Output file isn't where you expected | `-b docx` writes to `build/`, `-M docx` / `make docx` write to `build/docx/`. With `docx_documents` set, the names come from `targetname`. |
| Figure numbers are wrong or all `1` | The fields haven't been refreshed. In Word, press Ctrl+A then F9. |
