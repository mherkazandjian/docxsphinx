# OOXML fragments and python-docx quirks

python-docx (>= 1.0) covers paragraphs, runs, tables, pictures and styles. It has **no public
API** for hyperlinks, bookmarks, fields, footnotes, math or content controls. docxsphinx builds
those as raw XML in `src/docxsphinx/_docx_helpers.py` using `docx.oxml.OxmlElement` and
`docx.oxml.ns.qn`:

```python
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

r = OxmlElement('w:r')           # <w:r>
r.set(qn('w:id'), '1')           # qn() gives the Clark-notation name for a prefixed tag
paragraph._p.append(r)           # paragraph._p is the underlying <w:p> lxml element
```

## External hyperlink (`add_hyperlink`)

```xml
<w:hyperlink r:id="rId7" w:history="1">
  <w:r>
    <w:rPr><w:color w:val="0000EE"/><w:u w:val="single"/><w:b/></w:rPr>
    <w:t xml:space="preserve">link text</w:t>
  </w:r>
  <!-- one <w:r> per StyledChunk, so bold / italic / strike inside the link survive -->
</w:hyperlink>
```

The URL is registered as an external relationship first:
`r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)`. Colour and underline go
on each run because templates often have no `Hyperlink` character style.

## Internal hyperlink and bookmark (`add_internal_hyperlink`, `add_bookmark`)

```xml
<w:hyperlink w:anchor="my-section" w:history="1"> <w:r>…</w:r> </w:hyperlink>

<w:bookmarkStart w:id="3" w:name="my-section"/>
<w:bookmarkEnd w:id="3"/>
```

Internal links need no relationship. `w:anchor` must equal a `w:bookmarkStart/@w:name`. The
bookmarks have zero width (start and end are adjacent). Their `w:id` values are integers that
must be unique across the document; they come from the translator's `_bookmark_counter`, and
`_emitted_bookmarks` stops a name being emitted twice. Section ids get bookmarks in
`visit_title`, figure ids in `visit_caption`.

## SEQ field for figure numbers (`add_seq_field`)

```xml
<w:fldSimple w:instr=" SEQ Figure \* ARABIC ">
  <w:r><w:t>1</w:t></w:r>      <!-- placeholder shown until Word updates fields -->
</w:fldSimple>
```

The writer fills in the placeholder from `_figure_counter`, so LibreOffice and other viewers that
don't recompute fields still show the right number. `paragraph.text` does not descend into
`w:fldSimple`; to read it, look at its `<w:t>` directly.

## Footnotes (`ensure_footnotes_part`, `add_footnote`, `add_footnote_reference`)

A separate part, `/word/footnotes.xml`, is created on first use and linked from the main
document part:

```python
element = parse_xml(_FOOTNOTES_BOILERPLATE)   # separator (id -1) + continuationSeparator (id 0)
part = XmlPart(PackURI('/word/footnotes.xml'), CT.WML_FOOTNOTES, element, doc.part.package)
doc.part.relate_to(part, RT.FOOTNOTES)        # internal relationship
```

Calling it again finds the existing relationship and returns the same root. User footnote ids
start at 1. Body, inside the footnotes part:

```xml
<w:footnote w:id="1">
  <w:p>
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> </w:t></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">body chunk</w:t></w:r>
  </w:p>
</w:footnote>
```

Reference, in the document body:

```xml
<w:r>
  <w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr>
  <w:footnoteReference w:id="1"/>
</w:r>
```

docutils footnote ids are mapped to OOXML ids through `_footnote_ooxml_id`, so a reference and its
body get the same number whichever is visited first. Footnote labels are SkipNode'd, because
Word numbers footnotes itself.

## Math (`latex_to_omml`)

pandoc runs once for each distinct `(latex, display)` pair (`lru_cache` over bytes):
`pandoc -f latex -t docx` on `$…$` or `\[…\]`. The `<m:oMath>` (inline) or `<m:oMathPara>`
(display) element is then cut out of the resulting `word/document.xml`. Its root gets an
`xmlns:m` declaration and it is re-parsed on every call, because an lxml element can only have
one parent. Inline math is appended to `current_paragraph._p`; display math goes into a new
paragraph. A `None` return (no pandoc, or a conversion error) means the caller falls back to
LaTeX text in Consolas.

## Table row spans (vMerge)

```xml
<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>…content…</w:tc>   <!-- top cell -->
<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>                      <!-- each continuation -->
```

docutils omits `<entry>` nodes for cells covered by a row span, but Word needs a continuation
`w:tc` in every spanned row. The writer records spans in `current_state.rowspan_until` (to skip
those columns while placing later entries) and `pending_vmerges`, then merges them in
`depart_table` with python-docx's `cell.merge()`. Column spans (`morecols`) are merged straight
away.

## Run formatting

| Effect | python-docx | XML |
|---|---|---|
| Bold / italic | `run.bold`, `run.italic` | `<w:b/>`, `<w:i/>` |
| Strike | `run.font.strike = True` | `<w:strike/>` |
| Sub / superscript | `run.font.subscript` / `.superscript` | `<w:vertAlign w:val="subscript"/>` |
| Line break inside a paragraph | `run.add_break()` | `<w:br/>` |
| Monospace | `run.font.name = 'Consolas'` (code blocks only) | `<w:rFonts …/>` |

## python-docx quirks

- `table.rows[i].cells[j]` collapses merged slots into the top cell's object. To see each row's
  real cells, iterate `tr.findall(qn('w:tc'))` on `table._tbl.findall(qn('w:tr'))`.
- `run.text` is `'\n'` for a run that holds only a break. Filter with `r.text and r.text.strip()`
  when counting text runs.
- `paragraph.text` / `paragraph.runs` only see direct `<w:r>` children. Text inside
  `w:hyperlink` or `w:fldSimple` is invisible; use `paragraph._p.findall(qn('w:hyperlink'))`.
- `parse_xml()` rejects a `str` that begins with an `<?xml …?>` declaration. Drop the
  declaration or pass bytes.
- `Document('x.dotx')` raises because the main part's content type is
  `…template.main+xml`. `DocxWriter._open_template` rewrites it to `…document.main+xml` in an
  in-memory copy of the zip.
- `Document.add_picture()` always starts a new paragraph at the document root. Inside a cell or
  a running paragraph, call `run.add_picture()` on a run at the current location.
- `run.add_picture()` raises `docx.image.exceptions.UnrecognizedImageError`, a plain `Exception`,
  for SVG / WebP. It does not raise `OSError` or `ValueError`.
- `add_table(..., style=…)` only takes effect for tables at the document root. Nested tables
  get no style.
- `docx.table._Cell` is private but the only handle for `isinstance` checks. Import it through
  `docxsphinx._compat._Cell`, which has a fallback stub.
- `styles.get_style_id(name, WD_STYLE_TYPE.X)` raises `KeyError` when a style is missing. That's
  the hook the writer uses to warn and fall back to no style.
