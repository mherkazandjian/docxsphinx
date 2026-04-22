# Round-trip findings

Phase A deliverable of the `docxsphinx` reverse-pipeline work
(see `.claude/plans/now-review-the-repo-nested-corbato.md`). Produced by
running `tools/roundtrip.py` on the committed corpus and, optionally, a
private corpus of real Word-authored documents.

## Method

For each `.docx`:

1. Convert to Markdown via `pandoc -f docx -t gfm`.
2. Feed that Markdown through docxsphinx's forward path inside a minimal
   on-the-fly Sphinx project.
3. Extract `word/document.xml` from both the original and the
   round-tripped `.docx`.
4. Tally paragraph styles, run styles, run-property flags, and the full
   per-tag element census.

The diff surfaces idiomatic patterns in the originals that docxsphinx's
forward path should start emitting — and patterns we emit that idiomatic
Word documents don't use.

Reproduce locally:

```
make roundtrip                      # against examples/corpus_samples/
make roundtrip CORPUS=corpus        # against your private corpus
```

Reports land in `examples/corpus/reports/` (gitignored).

## Current corpus

- `examples/corpus_samples/pandoc_md_basic.docx` — pandoc's docx output
  from `examples/md_basic/source/index.md`. Represents pandoc's idiomatic
  Word emission of a simple Markdown document.

Additional reference fixtures (LibreOffice-authored, Word-authored,
Google-Docs-exported) should be dropped in `examples/corpus/` and the
harness re-run. The results below are from pandoc-only; expanding the
corpus will sharpen them.

## Findings — ranked by expected-impact

### 1. Numbered/bulleted lists are not real Word lists

**Severity**: high (visible to end users as soon as they edit in Word).

Pandoc's (and MS Word's) idiomatic list representation attaches a
numbering definition to each list-item paragraph:

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Compact"/>
    <w:numPr>
      <w:ilvl w:val="0"/>
      <w:numId w:val="1"/>
    </w:numPr>
  </w:pPr>
  …
</w:p>
```

docxsphinx emits paragraph-level style names (`ListBullet`, `ListNumber`)
but **no `w:numPr`/`w:numId`/`w:ilvl`** at all. The result looks like a
list in Word but behaves as a sequence of independently styled
paragraphs — editing the text breaks the visual numbering; indent/outdent
buttons do nothing; the list cannot be restarted or continued.

Round-trip signal: `numPr`, `numId`, `ilvl` all at 6 occurrences in the
original and 0 in our output.

**Action**: emit `<w:numPr>` with a `numId` from a generated numbering
part (`word/numbering.xml`) on each list-item paragraph. Requires adding
a new OPC part similar to how footnotes are wired. Belongs in
`_docx_helpers.py`.

### 2. Inline code runs have no character style

**Severity**: medium (code looks visually similar but has no semantic
marker).

Pandoc applies the built-in Word style `VerbatimChar` as a run-level
`<w:rStyle>` to every inline `<code>` span:

```xml
<w:r>
  <w:rPr><w:rStyle w:val="VerbatimChar"/></w:rPr>
  <w:t>code literal</w:t>
</w:r>
```

docxsphinx emits inline `literal` nodes as plain runs. No `rStyle` at
all. Five inline-code spans in the source → zero `rStyle` elements in
our output.

**Action**: set `<w:rStyle w:val="VerbatimChar">` on runs emitted from
`visit_literal`. Possibly fall back to setting `font.name = 'Consolas'`
manually if the template lacks the style (mirrors the current
`ensure_style` pattern).

### 3. Body-text paragraphs use only `Normal` style

**Severity**: medium (documents look unprofessional in Word next to
tool-generated reference output).

Pandoc differentiates three kinds of body paragraph:

- `FirstParagraph` — the first paragraph under each heading (4× here).
- `BodyText` — a subsequent paragraph (1× here).
- `Compact` — a paragraph inside a list item (6× here).

docxsphinx emits all of these as `Normal` (via python-docx's default).
Consequence: Word's paragraph-spacing rules that apply to these specific
styles never fire, and the document spacing looks subtly off compared
to any document authored in Word or emitted by pandoc.

**Action**: track paragraph context in the visitor and select:

- `FirstParagraph` for the first `<p>` under each `<section>` / `<title>`.
- `BodyText` for subsequent paragraphs.
- `Compact` for list-item child paragraphs.

Fallback to `Normal` if the template doesn't provide the style.

### 4. Bold/italic without the complex-script siblings

**Severity**: low (visual parity; matters only for mixed-script content).

Word's OOXML expects `<w:b/>` **and** `<w:bCs/>` (complex-script bold)
together, so bold renders in both Latin and non-Latin scripts. Same for
`<w:i/>` + `<w:iCs/>`. Pandoc emits both; we emit only the non-complex
forms.

Round-trip: 2 `w:b` vs 0 `w:bCs`; 2 `w:i` vs 0 `w:iCs`.

**Action**: when setting `run.font.bold = True` / `italic = True`,
additionally append `<w:bCs/>` / `<w:iCs/>` to the run's `<w:rPr>`.
Small diff in `writer.py::add_text`.

### 5. Section properties present but minimal

**Severity**: low (cosmetic — but docxsphinx-only tags that aren't in
originals suggest we should either match Word's defaults more closely or
elide these when not needed).

docxsphinx emits a section-level `<w:sectPr>` containing `<w:pgSz>`,
`<w:pgMar>`, `<w:cols>`, `<w:docGrid>` (i.e. page size, margins,
columns, grid). Pandoc's output in our fixture omits these. Whether
this is a real problem depends on whether Word renders our pages
differently from an unadorned sectPr — likely not, but worth
investigating on a real-world corpus.

**Action**: deferred — add to the backlog for a broader-corpus run.

## Patterns *docxsphinx* emits that originals don't

- `ListBullet`, `ListNumber` paragraph styles — will disappear once
  finding #1 is addressed (we'll drive lists via `numPr` instead).

## Recommended follow-up order (separate PRs, outside the reverse plan)

1. **Real numbered/bulleted lists** via `numbering.xml` part + `w:numPr`.
   High-impact, medium-complexity.
2. **`VerbatimChar` run style on inline code**. Trivial-complexity,
   visible quality win.
3. **`FirstParagraph` / `BodyText` / `Compact` paragraph-style
   differentiation**. Medium-complexity, moderate quality win.
4. **`w:bCs` / `w:iCs` complex-script variants**. Trivial; bundle with
   any `add_text` rework.
5. Re-run the harness after each lands; expect the aggregate report to
   shrink in the corresponding row.

## Expand the corpus before drawing stronger conclusions

A single pandoc-authored fixture is a sample size of one. Before making
any of the above changes permanent, expand the corpus with:

- One or two documents authored directly in MS Word (local user task —
  drop them in `examples/corpus/`).
- A LibreOffice-authored document (representative of open-source authoring).
- A Google Docs export (representative of SaaS authoring).

The same findings against a broader corpus become confidently actionable.
