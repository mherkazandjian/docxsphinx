"""Sphinx builder that emits one or more OpenXML (.docx) documents.

Originally based on sphinxcontrib-docxbuilder (copyright 2010 shimizukawa,
BSD licence).
"""
from __future__ import annotations

import logging as _stdlib_logging
from collections.abc import Iterable
from pathlib import Path
from typing import NamedTuple

from docutils import nodes
from docutils.io import StringOutput
from sphinx import addnodes
from sphinx.builders import Builder
from sphinx.util import logging

from docxsphinx._compat import (
    bold,
    darkgreen,
    ensuredir,
    inline_all_toctrees,
    os_path,
)
from docxsphinx.writer import DocxWriter

logger = logging.getLogger(__name__)

_DEBUG_HANDLER_NAME = '_docxsphinx_debug_file'


class DocxDocumentEntry(NamedTuple):
    """One entry in ``docx_documents`` — a single output target.

    Fields mirror a subset of Sphinx's ``latex_documents``:
    - ``startdoc``: source docname (no suffix) serving as the root of
      this output's toctree traversal.
    - ``targetname``: output filename, with or without the ``.docx``
      suffix; the builder appends it if missing.
    - ``template``: optional per-output ``.docx`` / ``.dotx`` template
      (path relative to ``srcdir`` or absolute). ``None`` falls back to
      the project-level ``docx_template``.
    - ``toctree_only``: when ``True``, emit only the content reached by
      ``startdoc``'s toctrees — dropping the master's own prose sections.
      Useful when the master doc is a cover page / index only.
    """
    startdoc: str
    targetname: str
    template: str | None = None
    toctree_only: bool = False


class DocxBuilder(Builder):
    name = 'docx'
    format = 'docx'
    out_suffix = '.docx'

    def init(self) -> None:
        self._install_debug_log_handler()

    def _install_debug_log_handler(self) -> None:
        """Attach a DEBUG-level FileHandler if ``docx_debug_log`` is set.

        ``docx_debug_log`` is a boolean or a path. When truthy it writes
        visitor-level trace output to ``<outdir>/docx.log`` (the default)
        or to the given path. Without this config value the builder is
        silent beyond Sphinx's own logging.
        """
        target = getattr(self.config, 'docx_debug_log', None)
        if not target:
            return
        writer_logger = _stdlib_logging.getLogger('docxsphinx.writer')
        if any(getattr(h, 'name', None) == _DEBUG_HANDLER_NAME for h in writer_logger.handlers):
            return
        path = Path(self.outdir) / 'docx.log' if target is True else Path(target)
        if not path.is_absolute():
            path = Path(self.outdir) / path
        path.parent.mkdir(parents=True, exist_ok=True)
        handler = _stdlib_logging.FileHandler(path, mode='w')
        handler.set_name(_DEBUG_HANDLER_NAME)
        handler.setLevel(_stdlib_logging.DEBUG)
        handler.setFormatter(_stdlib_logging.Formatter(
            '%(asctime)-15s  %(name)s  %(message)s',
        ))
        writer_logger.addHandler(handler)
        writer_logger.setLevel(_stdlib_logging.DEBUG)

    def get_outdated_docs(self) -> str:
        return 'pass'

    def get_target_uri(self, docname: str, typ: str | None = None) -> str:
        return ''

    def fix_refuris(self, tree: nodes.Element) -> None:
        """Strip double-anchor refuris down to a single in-document anchor."""
        fname = self.config.master_doc + self.out_suffix
        for refnode in tree.traverse(nodes.reference):
            if 'refuri' not in refnode:
                continue
            refuri = refnode['refuri']
            hashindex = refuri.find('#')
            if hashindex < 0:
                continue
            hashindex = refuri.find('#', hashindex + 1)
            if hashindex >= 0:
                refnode['refuri'] = fname + refuri[hashindex:]

    def prepare_writing(self, docnames) -> None:
        # No longer a single-writer builder: one writer is instantiated per
        # output entry in ``write()`` so that each can carry its own
        # template and start from a clean ``Document``. Preserved as a
        # no-op for the Sphinx builder contract.
        pass

    def assemble_doctree(
        self, startdoc: str, toctree_only: bool = False,
    ) -> nodes.Element:
        """Assemble the single doctree for one output entry.

        When ``toctree_only`` is true, strip ``startdoc``'s own
        paragraph / section content and keep only the master title and
        the contents reachable through its toctree directives — mirroring
        Sphinx's ``latex_documents`` semantics.
        """
        tree = self.env.get_doctree(startdoc)
        if toctree_only:
            tree = _toctree_only_shell(tree)
        tree = inline_all_toctrees(self, set(), startdoc, tree, darkgreen, [startdoc])
        tree['docname'] = startdoc
        self.env.resolve_references(tree, startdoc, self)
        self.fix_refuris(tree)
        return tree

    def _document_entries(self) -> list[DocxDocumentEntry]:
        """Return the list of output entries to build.

        Back-compat: if ``docx_documents`` is empty/unset, synthesise a
        single entry from ``project`` + ``version`` + ``master_doc`` —
        reproducing the pre-2.1 single-output behaviour exactly.
        """
        configured = getattr(self.config, 'docx_documents', None) or []
        entries: list[DocxDocumentEntry] = []
        for raw in configured:
            entries.append(DocxBuilder._coerce_entry(raw))
        if entries:
            return entries

        default_name = f'{self.config.project}-{self.config.version}'
        return [DocxDocumentEntry(
            startdoc=self.config.master_doc,
            targetname=default_name,
            template=None,
            toctree_only=False,
        )]

    @staticmethod
    def _coerce_entry(raw: Iterable) -> DocxDocumentEntry:
        """Accept either a ``DocxDocumentEntry`` or a raw tuple/list
        ``(startdoc, targetname[, template[, toctree_only]])``."""
        if isinstance(raw, DocxDocumentEntry):
            return raw
        parts = list(raw)
        if len(parts) < 2:
            raise ValueError(
                'docx_documents entries need at least (startdoc, targetname); '
                f'got {raw!r}'
            )
        startdoc = parts[0]
        targetname = parts[1]
        template = parts[2] if len(parts) >= 3 else None
        toctree_only = bool(parts[3]) if len(parts) >= 4 else False
        return DocxDocumentEntry(startdoc, targetname, template, toctree_only)

    def write(self, *ignored) -> None:
        logger.info(bold('preparing documents... '), nonl=True)
        self.prepare_writing(self.env.all_docs)
        logger.info('done')

        entries = self._document_entries()
        for entry in entries:
            logger.info(
                bold('assembling document %r... '), entry.targetname, nonl=True,
            )
            doctree = self.assemble_doctree(entry.startdoc, entry.toctree_only)
            logger.info('')
            logger.info(bold('writing %r... '), entry.targetname, nonl=True)
            writer = DocxWriter(self, template_override=entry.template)
            self._write_one(writer, entry.targetname, doctree)
            logger.info('done')

    def _write_one(
        self, writer: DocxWriter, targetname: str, doctree: nodes.Element,
    ) -> None:
        destination = StringOutput(encoding='utf-8')
        writer.write(doctree, destination)
        name = targetname
        if not name.endswith(self.out_suffix):
            name += self.out_suffix
        outfilename = Path(self.outdir) / os_path(name)
        ensuredir(str(outfilename.parent))
        try:
            writer.save(str(outfilename))
        except OSError as err:
            logger.warning('error writing file %s: %s', outfilename, err)

    # Kept for backwards-compatibility with any external caller that used
    # the old single-writer API; delegates to ``_write_one`` using a
    # freshly-instantiated writer.
    def write_doc(self, docname: str, doctree: nodes.Element) -> None:
        writer = DocxWriter(self)
        self._write_one(writer, docname, doctree)

    def finish(self) -> None:
        pass


def _toctree_only_shell(master_tree: nodes.document) -> nodes.document:
    """Return a shallow copy of ``master_tree`` containing only the
    master title plus the project's toctrees.

    Mirrors the ``toctree_only`` behaviour of Sphinx's LaTeX builder:
    the master doc is treated as a cover/index whose own prose content
    is dropped, so the output begins at the first toctree-included
    chapter.
    """
    def _is_keeper(n) -> bool:
        """True for nodes we keep inside / alongside the master title."""
        if isinstance(n, (nodes.title, addnodes.toctree)):
            return True
        return isinstance(n, nodes.compound) and any(
            isinstance(gc, addnodes.toctree) for gc in n.children
        )

    new_tree = master_tree.copy()
    new_tree.clear()
    # The first child is usually a section wrapping the master title —
    # preserve it (Word still wants a top-level Heading 1) but strip
    # every child of that section except the title and toctrees.
    for node in master_tree.children:
        if isinstance(node, nodes.section):
            wrapper = node.copy()
            for child in node.children:
                if _is_keeper(child):
                    wrapper += child.deepcopy()
            new_tree += wrapper
        elif _is_keeper(node):
            new_tree += node.deepcopy()
    return new_tree
