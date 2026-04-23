"""Deterministic text fingerprint of a ``.docx``'s content.

Extracts the structural signal we care about (paragraph counts, style
distribution, heading texts in order, table shapes, hyperlinks, bookmark
names, footnote count, inline-shape count, element-tag census) and
formats it as a stable, human-readable text block. Used by the
golden-file tests in ``tests/test_golden.py`` — regressions that silently
drop content (e.g. a visitor starts ``SkipNode``-ing) produce a visible
diff in the committed fingerprint.

Deliberately omits anything that varies run-to-run:

- relationship IDs (``rId1``, ``rId2``, …)
- run-level colour values (vary with pygments version)
- section-properties page-geometry defaults
- ``<w:rsid*>`` revision-save IDs
- timestamps / bookmark IDs (the numeric ones — names are kept)
"""
from __future__ import annotations

import io
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

W = '{' + WML_NS + '}'
R = '{' + REL_NS + '}'


def _local(tag: str) -> str:
    if tag.startswith(W):
        return 'w:' + tag[len(W):]
    if tag.startswith('{' + MATH_NS + '}'):
        return 'm:' + tag[len('{' + MATH_NS + '}'):]
    return tag


def _text_of(el: etree._Element) -> str:
    """Concatenate all <w:t> descendants into one string."""
    parts = []
    for t in el.iter(W + 't'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts).strip()


def _collect_hyperlinks(doc_xml: etree._Element, rels: dict[str, str]) -> list[str]:
    """Return list of normalised hyperlink lines: 'external <url> | <text>' or
    'anchor <name> | <text>'. Sorted for determinism."""
    lines = []
    for hl in doc_xml.iter(W + 'hyperlink'):
        text = _text_of(hl)
        rid = hl.get(R + 'id')
        anchor = hl.get(W + 'anchor')
        if rid and rid in rels:
            lines.append(f'  external  {rels[rid]:<50}  {text}')
        elif anchor:
            lines.append(f'  anchor    {anchor:<50}  {text}')
    lines.sort()
    return lines


def _load_rels(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map relationship id → target URL/partname from document.xml.rels."""
    try:
        xml = zf.read('word/_rels/document.xml.rels')
    except KeyError:
        return {}
    root = etree.fromstring(xml)
    # relationship namespace — different from the 'R' above
    rels_ns = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    return {
        rel.get('Id'): rel.get('Target', '')
        for rel in root.iter(rels_ns + 'Relationship')
    }


def _collect_tables(doc_xml: etree._Element) -> list[str]:
    """Return list of table shapes: '<rows>x<cols> | <first-cell-text truncated>'."""
    lines = []
    for tbl in doc_xml.iter(W + 'tbl'):
        rows = tbl.findall(W + 'tr')
        n_rows = len(rows)
        n_cols = max((len(r.findall(W + 'tc')) for r in rows), default=0)
        first_cell = ''
        if rows and rows[0].findall(W + 'tc'):
            first_cell = _text_of(rows[0].findall(W + 'tc')[0])[:40]
        lines.append(f'  {n_rows}x{n_cols}  {first_cell}')
    return lines


def _collect_bookmarks(doc_xml: etree._Element) -> list[str]:
    names = set()
    for bm in doc_xml.iter(W + 'bookmarkStart'):
        name = bm.get(W + 'name')
        if name and not name.startswith('_'):
            names.add(name)
    return sorted(names)


def _collect_headings(doc_xml: etree._Element) -> list[str]:
    """Return list of '<style>  <text>' for paragraphs whose style starts
    with 'Heading' or is 'Title'. Order preserved (document order)."""
    lines = []
    for p in doc_xml.iter(W + 'p'):
        style_el = p.find(W + 'pPr/' + W + 'pStyle')
        if style_el is None:
            continue
        style = style_el.get(W + 'val', '')
        if style == 'Title' or style.startswith('Heading'):
            text = _text_of(p)[:80]
            lines.append(f'  {style:<12}  {text}')
    return lines


def _collect_style_counts(doc_xml: etree._Element, tag: str) -> Counter:
    """Count <w:pStyle> or <w:rStyle> occurrences by val."""
    counts: Counter = Counter()
    for el in doc_xml.iter(W + tag):
        counts[el.get(W + 'val', '')] += 1
    return counts


def _collect_element_census(doc_xml: etree._Element) -> Counter:
    """Per-tag count (local names only) across the whole document."""
    counts: Counter = Counter()
    for el in doc_xml.iter():
        tag = el.tag
        if not isinstance(tag, str):
            continue
        counts[_local(tag)] += 1
    return counts


def _count_footnotes(zf: zipfile.ZipFile) -> int:
    """Count *user* footnotes (excluding the two mandatory separator entries)."""
    try:
        xml = zf.read('word/footnotes.xml')
    except KeyError:
        return 0
    root = etree.fromstring(xml)
    user = 0
    for fn in root.iter(W + 'footnote'):
        wid = fn.get(W + 'id')
        wtype = fn.get(W + 'type')
        if wtype in ('separator', 'continuationSeparator'):
            continue
        if wid is not None:
            try:
                if int(wid) > 0:
                    user += 1
            except ValueError:
                pass
    return user


def _count_inline_shapes(doc_xml: etree._Element) -> int:
    return len(list(doc_xml.iter(W + 'drawing')))


_TOP_TAG_CENSUS = 25
_REDACT_KEYS = {
    'sectPr', 'pgSz', 'pgMar', 'cols', 'docGrid', 'rsidR', 'rsidRPr',
    'rsidRDefault', 'rsidP', 'color', 'sz', 'szCs',
    'lang', 'rFonts', 'pPrChange', 'rPrChange', 'proofErr', 'bookmarkEnd',
}


def fingerprint(docx_path: Path) -> str:
    """Extract a deterministic text fingerprint from a .docx file."""
    with zipfile.ZipFile(docx_path) as zf:
        doc_xml = etree.fromstring(zf.read('word/document.xml'))
        rels = _load_rels(zf)
        footnote_count = _count_footnotes(zf)

    out = io.StringIO()
    paragraph_count = len(list(doc_xml.iter(W + 'p')))
    table_count = len(list(doc_xml.iter(W + 'tbl')))
    inline_shapes = _count_inline_shapes(doc_xml)
    hyperlink_lines = _collect_hyperlinks(doc_xml, rels)
    bookmark_names = _collect_bookmarks(doc_xml)
    heading_lines = _collect_headings(doc_xml)
    pstyle_counts = _collect_style_counts(doc_xml, 'pStyle')
    rstyle_counts = _collect_style_counts(doc_xml, 'rStyle')
    element_census = _collect_element_census(doc_xml)

    print(f'# docxsphinx fingerprint for {docx_path.name}', file=out)
    print('# format-version: 1', file=out)
    print(file=out)
    print(f'paragraphs        {paragraph_count}', file=out)
    print(f'tables            {table_count}', file=out)
    print(f'inline_shapes     {inline_shapes}', file=out)
    print(f'hyperlinks        {len(hyperlink_lines)}', file=out)
    print(f'bookmarks         {len(bookmark_names)}', file=out)
    print(f'footnotes         {footnote_count}', file=out)
    print(file=out)

    print('# heading texts (in document order)', file=out)
    if heading_lines:
        for line in heading_lines:
            print(line, file=out)
    else:
        print('  (none)', file=out)
    print(file=out)

    print('# paragraph styles (sorted)', file=out)
    if pstyle_counts:
        for name in sorted(pstyle_counts):
            print(f'  {name:<28}  {pstyle_counts[name]}', file=out)
    else:
        print('  (none)', file=out)
    print(file=out)

    print('# run styles (sorted)', file=out)
    if rstyle_counts:
        for name in sorted(rstyle_counts):
            print(f'  {name:<28}  {rstyle_counts[name]}', file=out)
    else:
        print('  (none)', file=out)
    print(file=out)

    print('# table shapes (document order)', file=out)
    if table_count:
        for line in _collect_tables(doc_xml):
            print(line, file=out)
    else:
        print('  (none)', file=out)
    print(file=out)

    print('# hyperlinks (sorted)', file=out)
    if hyperlink_lines:
        for line in hyperlink_lines:
            print(line, file=out)
    else:
        print('  (none)', file=out)
    print(file=out)

    print('# bookmark names (sorted)', file=out)
    if bookmark_names:
        for name in bookmark_names:
            print(f'  {name}', file=out)
    else:
        print('  (none)', file=out)
    print(file=out)

    print(f'# element tag census (top {_TOP_TAG_CENSUS} — excluding noisy infra tags)', file=out)
    filtered = [
        (tag, n) for tag, n in element_census.most_common()
        if tag not in {f'w:{k}' for k in _REDACT_KEYS}
    ][:_TOP_TAG_CENSUS]
    for tag, n in filtered:
        print(f'  {tag:<28}  {n}', file=out)

    return out.getvalue()


if __name__ == '__main__':
    import sys
    for arg in sys.argv[1:]:
        print(fingerprint(Path(arg)))
