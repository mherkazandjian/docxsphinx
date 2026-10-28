# docxsphinx feature matrix (main, after 2.1.0)

What each source construct becomes in the `.docx`, as of the main branch. The 2.1.0 release on
PyPI behaves differently for the constructs listed in the last section. When a construct isn't
listed, look up its node type with `sphinx-build -b pseudoxml`: a node type docxsphinx has no
renderer for keeps its text, unformatted, and logs one `docx.unknown_node` warning.

## Renders

| Construct (RST / MyST) | In Word |
|---|---|
| Section titles, `#`…`######` | `Heading 1`…`Heading N` by depth. Depth carries through nested toctrees. Each section id gets a bookmark. |
| Paragraphs | Normal paragraphs |
| `**bold**`, `*italic*`, `:sub:` / `:sup:` | Run bold / italic / subscript / superscript |
| `~~strike~~` (MyST `strikethrough` extension) | Run strikethrough |
| Bullet / numbered lists, nested and mixed | `List Bullet[ N]` / `List Number[ N]` by depth |
| `- [ ]` / `- [x]` (MyST `tasklist`) | `☐` / `☒` glyph at the start of a bullet paragraph (a visible character, not a clickable checkbox) |
| Definition lists, `.. glossary::` | 2-column table: bold term, then definition. Inline formatting in the term is flattened; RST classifiers (`term : classifier`) are dropped. |
| Field lists (`:name: body`), autodoc `:param:` / `:returns:` / `:raises:` | 2-column table: bold field name, then body. The parameter name in `:param x:` is bold. |
| Option lists (`-v, --verbose  text`) | 2-column table: flag signature, then description |
| Grid / simple tables, `list-table`, `csv-table`, GFM pipe tables | Native Word table with column and row spans merged. Style is `Grid Table 4`, or the one set by a `DocxTableStyle` comment. |
| `::` literal blocks, `code-block`, `literalinclude`, MyST fences | Pygments-coloured Consolas runs in a `Preformatted Text` paragraph |
| `code-block` with `:caption:` | A plain `Caption` paragraph (no "Figure N."), then the code. `:name:` gets a bookmark, so `:ref:` to the block works. |
| Bare doctest blocks (a paragraph starting `>>>`) | A code block, highlighted as a Python console session |
| `.. image::`, `![alt](x.png)`, `<img>` (MyST `html_image`) | Embedded picture. Sizes in px, cm, mm, in or pt; `:align:` left / center / right; alt text is kept. An image inside a paragraph stays inline. |
| `.. figure::` caption, `:name:` / `(label)=` | `Caption` paragraph reading "Figure N." from a Word SEQ field. The figure id gets a bookmark. |
| Figure legend (the paragraphs after the caption) | Normal paragraphs after the caption |
| `note`, `warning`, `tip`, `caution`, `important`, `hint`, `error`, `danger`, `attention`, generic `admonition` (RST or MyST `:::{note}`) | 1×1 table with a bold title line, then the body |
| `.. topic::`, `.. sidebar::`, `.. todo::` (needs `todo_include_todos = True`) | The same 1×1 box, titled from the directive's title (`Todo` for todos). A sidebar's subtitle is an italic line under the title. |
| `.. contents::` | The same box, titled "Contents" or the given title, holding a bulleted list of links to the section headings |
| `.. rubric::` | A bold paragraph (not a Word heading) |
| Line blocks (`\| line`) | One paragraph, a line break per line; nested blocks are indented |
| `versionadded`, `versionchanged`, `deprecated` | The body, starting with Sphinx's label ("Added in version 1.2: …", "Deprecated since version 2.0.") |
| `.. container::` | Its content in place, with no box or styling |
| Block quotes (indented text) | Normal paragraphs, no indentation or quote style |
| Block-quote attribution (`-- Author`) | A paragraph reading "— Author" |
| `.. productionlist::` | One Consolas paragraph, a line per production. Token names are bookmarked, so `:token:` links resolve. |
| `.. acks::` | Its bullet list |
| `.. seealso::` | Its paragraphs, with no box and no "See also" label |
| External links | Word hyperlink. Bold, italic and strike inside the link text are kept. |
| `:ref:`, `[text](#anchor)`, `(label)=` targets | Internal hyperlink to the section or figure bookmark |
| `:numref:` (needs `numfig = True`) | Internal hyperlink whose text is Sphinx's label ("Fig. 1") |
| Footnotes and citations | Native Word footnotes; formatting inside the body is kept |
| `:math:`, `.. math::`, `$…$`, `$$…$$`, `\begin{align}` (MyST `dollarmath` / `amsmath`) | Editable Word equations (OMML) when pandoc is on `PATH`, otherwise raw LaTeX in Consolas |
| autodoc (`automodule`, `autoclass`, `autofunction`, …) | One paragraph per signature, name in bold, `[…]` around optional args, `→` before the return type. The module prefix is omitted. Docstring fields become a 2-column table. |
| `.. compound::`, `.. hlist::`, `.. centered::` | Their contents as normal paragraphs |
| Directives from other extensions whose node type docxsphinx doesn't know | Their text as plain paragraphs, without that directive's look. One `docx.unknown_node` warning per node type. |

## Rendered differently than you might expect

| Construct | What happens |
|---|---|
| Inline code (``` ``x`` ```, `` `x` ``, `:code:`) | Plain text in the surrounding font, not monospace |
| Image `:scale:` or `%` width | Ignored with a warning; the picture keeps its native size |
| Remote image URL | Skipped with a `failed to embed image` warning |
| SVG, WebP or any format other than PNG / JPEG / GIF / BMP / TIFF | Skipped with a `docx.image` warning; convert it to PNG |
| Figure number | Placeholder value until Word updates fields (Ctrl+A, then F9) |
| `.. raw::` | Dropped. The one exception is MyST's `<s>` / `</s>` markers, which turn on strikethrough. |
| Transition (`----`) | Nothing |
| Comments, `.. index::`, substitution definitions, `.. meta::` | No output, as expected |

The `docx.unknown_node` and `docx.image` warnings fail a `-W` build. Silence them with
`suppress_warnings = ['docx.unknown_node', 'docx.image']`, or `['docx']` for both.

## On 2.1.0 (current PyPI release)

These were dropped **silently, content and all**, and an SVG or other unsupported image
**aborted the build** (convert it to PNG first). Any node type without a renderer vanished
with no warning. Workarounds for 2.1.0:

| Construct | Node type | Alternative that renders on 2.1.0 |
|---|---|---|
| `.. contents::`, `.. topic::` | `topic` | Word's References → Table of Contents after the build; plain paragraphs for a topic |
| `.. sidebar::` | `sidebar` | An admonition (`.. note::` or a generic `.. admonition:: Title`) |
| `.. container::`, incl. every `code-block` with `:caption:` | `container` | Plain content; a `code-block` without `:caption:` |
| `.. rubric::` | `rubric` | A heading, or a bold paragraph |
| Line blocks (`\| line`) | `line_block` | Separate paragraphs, or a literal block |
| Bare doctest blocks (a paragraph starting `>>>`) | `doctest_block` | `.. code-block:: pycon` |
| `versionadded`, `versionchanged`, `deprecated` | `versionmodified` | `.. note::` / `.. warning::` with the version written out |
| `.. todo::` | `todo_node` | `.. note::` |
| Figure legend | `legend` | A normal paragraph after the figure |
| Block-quote attribution (`-- Author`) | `attribution` | A normal paragraph |
| `.. productionlist::`, `.. acks::` | `productionlist`, `acks` | A literal block, or a list |
| autodoc `:param x:` name | `literal_strong` | (none; the name is lost, the description stays) |

## Not available

No Word table of contents field, page breaks, headers, footers, section or column layout, or
clickable form-field checkboxes. Put fixed parts (headers, footers, cover page) in the template.
Add a TOC in Word after the build. It works because headings use the real `Heading N` styles.
