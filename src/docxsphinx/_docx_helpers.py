"""Low-level OOXML helpers for emitting constructs that python-docx does
not expose through its public API — hyperlinks, bookmarks, footnotes, etc.

These helpers manipulate the underlying ``lxml`` element tree via python-docx's
``docx.oxml`` module. They are intentionally kept free of writer-specific
state so they can be unit-tested in isolation.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# RGB hex for the hyperlink blue Word uses by default. Applied as an
# explicit run-property colour so the link shows up visibly even when the
# target template has no "Hyperlink" character style.
_HYPERLINK_COLOR_HEX = '0000EE'


StyledChunk = tuple[str, bool, bool, bool]
"""``(text, bold, italic, strike)`` — the minimum run-level formatting we
propagate through hyperlink text and footnote bodies."""


def _styled_run(
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    strike: bool = False,
    color_hex: str | None = None,
    underline: bool = False,
) -> OxmlElement:
    """Build a single ``<w:r>`` with the given inline formatting."""
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if color_hex is not None:
        color_el = OxmlElement('w:color')
        color_el.set(qn('w:val'), color_hex)
        rPr.append(color_el)
    if underline:
        u_el = OxmlElement('w:u')
        u_el.set(qn('w:val'), 'single')
        rPr.append(u_el)
    if bold:
        rPr.append(OxmlElement('w:b'))
    if italic:
        rPr.append(OxmlElement('w:i'))
    if strike:
        rPr.append(OxmlElement('w:strike'))
    # w:rPr is only emitted if it has any children.
    if len(rPr):
        run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    run.append(t)
    return run


def _normalise_text_input(
    text_or_chunks: str | list[StyledChunk],
) -> list[StyledChunk]:
    if isinstance(text_or_chunks, str):
        return [(text_or_chunks, False, False, False)]
    return text_or_chunks


def add_hyperlink(
    paragraph: Paragraph,
    url: str,
    text_or_chunks: str | list[StyledChunk],
) -> OxmlElement:
    """Append a clickable *external* hyperlink to ``paragraph``.

    ``text_or_chunks`` is either a plain string (single run) or a list of
    ``(text, bold, italic, strike)`` tuples describing per-run formatting
    within the link anchor (preserved from the docutils reference node's
    inline structure). The hyperlink's blue+underline base styling is
    applied on top of the per-chunk bold/italic/strike flags.
    """
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    hyperlink.set(qn('w:history'), '1')
    for text, bold, italic, strike in _normalise_text_input(text_or_chunks):
        hyperlink.append(_styled_run(
            text, bold=bold, italic=italic, strike=strike,
            color_hex=_HYPERLINK_COLOR_HEX, underline=True,
        ))

    paragraph._p.append(hyperlink)
    return hyperlink


def add_internal_hyperlink(
    paragraph: Paragraph,
    anchor: str,
    text_or_chunks: str | list[StyledChunk],
) -> OxmlElement:
    """Append a clickable *internal* hyperlink pointing at a named bookmark.

    See :func:`add_hyperlink` for the ``text_or_chunks`` parameter.
    """
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), anchor)
    hyperlink.set(qn('w:history'), '1')
    for text, bold, italic, strike in _normalise_text_input(text_or_chunks):
        hyperlink.append(_styled_run(
            text, bold=bold, italic=italic, strike=strike,
            color_hex=_HYPERLINK_COLOR_HEX, underline=True,
        ))

    paragraph._p.append(hyperlink)
    return hyperlink


def add_bookmark(
    paragraph: Paragraph, name: str, bookmark_id: int,
) -> tuple[OxmlElement, OxmlElement]:
    """Emit a zero-width bookmark (``start`` immediately followed by
    ``end``) at the end of ``paragraph``.

    Bookmarks are the target mechanism for internal hyperlinks via
    ``w:hyperlink/@w:anchor``. ``bookmark_id`` must be unique within
    the document.
    """
    bm_start = OxmlElement('w:bookmarkStart')
    bm_start.set(qn('w:id'), str(bookmark_id))
    bm_start.set(qn('w:name'), name)
    bm_end = OxmlElement('w:bookmarkEnd')
    bm_end.set(qn('w:id'), str(bookmark_id))
    paragraph._p.append(bm_start)
    paragraph._p.append(bm_end)
    return bm_start, bm_end


def add_seq_field(
    paragraph: Paragraph, name: str = 'Figure', initial_value: int | None = None,
) -> OxmlElement:
    """Append a Word ``SEQ`` field to ``paragraph`` for auto-numbering.

    ``name`` identifies the sequence (``'Figure'``, ``'Table'``, …). Each
    occurrence of ``SEQ name`` in a document auto-increments when Word
    updates fields (Ctrl-A → F9, or on file open for some installations).

    ``initial_value`` supplies the digit rendered inline *before* Word
    updates fields. That makes freshly-generated documents readable even
    in apps (LibreOffice, older viewers) that don't refresh SEQ fields.
    Defaults to ``1``.
    """
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), f' SEQ {name} \\* ARABIC ')
    run = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = str(initial_value if initial_value is not None else 1)
    run.append(t)
    fld.append(run)
    paragraph._p.append(fld)
    return fld


_FOOTNOTES_BOILERPLATE = (
    '<w:footnotes '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:footnote w:type="separator" w:id="-1">'
    '<w:p><w:r><w:separator/></w:r></w:p>'
    '</w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:r><w:continuationSeparator/></w:r></w:p>'
    '</w:footnote>'
    '</w:footnotes>'
)


def ensure_footnotes_part(document: Document) -> OxmlElement:
    """Return the document's ``<w:footnotes>`` root element, creating the
    ``word/footnotes.xml`` part and wiring its relationship on first call.

    Callers append ``<w:footnote w:id="…">`` children to the returned
    element via :func:`add_footnote`. python-docx has no public footnote
    API as of 1.2, so this function manipulates the underlying OPC package
    directly.
    """
    doc_part = document.part
    for rel in doc_part.rels.values():
        if rel.reltype == RT.FOOTNOTES:
            return rel.target_part.element  # type: ignore[attr-defined]

    element = parse_xml(_FOOTNOTES_BOILERPLATE)
    partname = PackURI('/word/footnotes.xml')
    part = XmlPart(partname, CT.WML_FOOTNOTES, element, doc_part.package)
    doc_part.relate_to(part, RT.FOOTNOTES)
    return element


def add_footnote(
    footnotes_root: OxmlElement,
    footnote_id: int,
    body_text_or_chunks: str | list[StyledChunk],
) -> OxmlElement:
    """Append a ``<w:footnote w:id="N">`` with a single paragraph containing
    the footnote marker + body to ``footnotes_root``.

    ``body_text_or_chunks`` is either a plain string (single run) or a
    list of ``(text, bold, italic, strike)`` styled chunks — the writer
    passes the latter when the footnote source contained inline
    formatting, so bold/italic/strike spans within a footnote body
    round-trip correctly.
    """
    footnote = OxmlElement('w:footnote')
    footnote.set(qn('w:id'), str(footnote_id))

    para = OxmlElement('w:p')

    # The canonical "footnoteRef + space + body" pattern: Word substitutes
    # the auto-number at render time for <w:footnoteRef/>.
    marker_run = OxmlElement('w:r')
    marker_rPr = OxmlElement('w:rPr')
    marker_style = OxmlElement('w:rStyle')
    marker_style.set(qn('w:val'), 'FootnoteReference')
    marker_rPr.append(marker_style)
    marker_run.append(marker_rPr)
    marker_run.append(OxmlElement('w:footnoteRef'))
    para.append(marker_run)

    chunks = _normalise_text_input(body_text_or_chunks)
    # Put the leading space as its own plain-text run so it isn't bold
    # just because the first chunk happens to be.
    para.append(_styled_run(' '))
    for text, bold, italic, strike in chunks:
        para.append(_styled_run(text, bold=bold, italic=italic, strike=strike))

    footnote.append(para)
    footnotes_root.append(footnote)
    return footnote


def add_footnote_reference(paragraph: Paragraph, footnote_id: int) -> OxmlElement:
    """Append a superscripted ``<w:footnoteReference w:id="N"/>`` run to
    ``paragraph``. Word renders this as the footnote's auto-number where it
    appears inline; clicking it jumps to the footnote pane."""
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    style_el = OxmlElement('w:rStyle')
    style_el.set(qn('w:val'), 'FootnoteReference')
    rPr.append(style_el)
    vert_align = OxmlElement('w:vertAlign')
    vert_align.set(qn('w:val'), 'superscript')
    rPr.append(vert_align)
    run.append(rPr)

    ref = OxmlElement('w:footnoteReference')
    ref.set(qn('w:id'), str(footnote_id))
    run.append(ref)

    paragraph._p.append(run)
    return run


__all__ = [
    'add_bookmark',
    'add_footnote',
    'add_footnote_reference',
    'add_hyperlink',
    'add_internal_hyperlink',
    'add_seq_field',
    'ensure_footnotes_part',
]
