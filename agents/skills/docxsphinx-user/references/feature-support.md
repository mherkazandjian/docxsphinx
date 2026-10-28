# docxsphinx feature matrix (2.1)

What each source construct becomes in the `.docx`. "Dropped" means the node **and everything
inside it** is left out, with no warning. When a construct isn't listed, look up its node type
(run `sphinx-build -b pseudoxml`): a node docxsphinx has no handler for is dropped.

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
| Field lists (`:name: body`), autodoc `:param:` / `:returns:` / `:raises:` | 2-column table: bold field name, then body |
| Option lists (`-v, --verbose  text`) | 2-column table: flag signature, then description |
| Grid / simple tables, `list-table`, `csv-table`, GFM pipe tables | Native Word table with column and row spans merged. Style is `Grid Table 4`, or the one set by a `DocxTableStyle` comment. |
| `::` literal blocks, `code-block`, `literalinclude`, MyST fences | Pygments-coloured Consolas runs in a `Preformatted Text` paragraph |
| `.. image::`, `![alt](x.png)`, `<img>` (MyST `html_image`) | Embedded picture. Sizes in px, cm, mm, in or pt; `:align:` left / center / right; alt text is kept. An image inside a paragraph stays inline. |
| `.. figure::` caption, `:name:` / `(label)=` | `Caption` paragraph reading "Figure N." from a Word SEQ field. The figure id gets a bookmark. |
| `note`, `warning`, `tip`, `caution`, `important`, `hint`, `error`, `danger`, `attention`, generic `admonition` (RST or MyST `:::{note}`) | 1×1 table with a bold title line, then the body |
| `.. seealso::` | Its paragraphs, with no box and no "See also" label |
| External links | Word hyperlink. Bold, italic and strike inside the link text are kept. |
| `:ref:`, `[text](#anchor)`, `(label)=` targets | Internal hyperlink to the section or figure bookmark |
| `:numref:` (needs `numfig = True`) | Internal hyperlink whose text is Sphinx's label ("Fig. 1") |
| Footnotes and citations | Native Word footnotes; formatting inside the body is kept |
| `:math:`, `.. math::`, `$…$`, `$$…$$`, `\begin{align}` (MyST `dollarmath` / `amsmath`) | Editable Word equations (OMML) when pandoc is on `PATH`, otherwise raw LaTeX in Consolas |
| autodoc (`automodule`, `autoclass`, `autofunction`, …) | One paragraph per signature, name in bold, `[…]` around optional args, `→` before the return type. The module prefix is omitted. Docstring fields become a 2-column table. |
| `.. compound::`, `.. hlist::`, `.. centered::` | Their contents as normal paragraphs |
| Block quotes (indented text) | Normal paragraphs, no indentation or quote style |

## Rendered differently than you might expect

| Construct | What happens |
|---|---|
| Inline code (``` ``x`` ```, `` `x` ``, `:code:`) | Plain text in the surrounding font, not monospace |
| Image `:scale:` or `%` width | Ignored with a warning; the picture keeps its native size |
| Remote image URL | Skipped with a `failed to embed image` warning |
| SVG, WebP or any format other than PNG / JPEG / GIF / BMP / TIFF | **The build aborts** (python-docx `UnrecognizedImageError`) |
| Figure number | Placeholder value until Word updates fields (Ctrl+A, then F9) |
| `.. raw::` | Dropped. The one exception is MyST's `<s>` / `</s>` markers, which turn on strikethrough. |
| Transition (`----`) | Nothing |
| Comments, `.. index::`, substitution definitions, `.. meta::` | No output, as expected |

## Dropped silently (content lost)

| Construct | Node type | Alternative that renders |
|---|---|---|
| `.. contents::`, `.. topic::` | `topic` | Word's References → Table of Contents after the build; plain paragraphs for a topic |
| `.. sidebar::` | `sidebar` | An admonition (`.. note::` or a generic `.. admonition:: Title`) |
| `.. container::`, and directives from other extensions that wrap content in one | `container` | Plain content, or an admonition |
| `.. rubric::` | `rubric` | A heading, or a bold paragraph |
| Line blocks (`\| line`) | `line_block` | Separate paragraphs, or a literal block |
| Bare doctest blocks (a paragraph starting `>>>`) | `doctest_block` | `.. code-block:: pycon` |
| `versionadded`, `versionchanged`, `deprecated` | `versionmodified` | `.. note::` / `.. warning::` with the version written out |
| `.. todo::` | `todo_node` | `.. note::` |
| Figure legend (the paragraphs after a figure's caption) | `legend` | A normal paragraph after the figure |
| Block-quote attribution (`-- Author`) | `attribution` | A normal paragraph |
| `.. productionlist::`, `.. acks::` | `productionlist`, `acks` | A literal block, or a list |
| Anything else with no visitor | (various) | Look up the node with `-b pseudoxml` |

## Not available

No Word table of contents field, page breaks, headers, footers, section or column layout, or
clickable form-field checkboxes. Put fixed parts (headers, footers, cover page) in the template.
Add a TOC in Word after the build. It works because headings use the real `Heading N` styles.
