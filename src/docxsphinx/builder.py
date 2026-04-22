"""Sphinx builder that emits a single OpenXML (.docx) document.

Originally based on sphinxcontrib-docxbuilder (copyright 2010 shimizukawa,
BSD licence).
"""
from __future__ import annotations

import logging as _stdlib_logging
from pathlib import Path

from docutils import nodes
from docutils.io import StringOutput
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
        self.writer = DocxWriter(self)

    def assemble_doctree(self) -> nodes.Element:
        master = self.config.master_doc
        tree = self.env.get_doctree(master)
        tree = inline_all_toctrees(self, set(), master, tree, darkgreen, [master])
        tree['docname'] = master
        self.env.resolve_references(tree, master, self)
        self.fix_refuris(tree)
        return tree

    def write(self, *ignored) -> None:
        logger.info(bold('preparing documents... '), nonl=True)
        self.prepare_writing(self.env.all_docs)
        logger.info('done')

        logger.info(bold('assembling single document... '), nonl=True)
        doctree = self.assemble_doctree()
        logger.info('')
        logger.info(bold('writing... '), nonl=True)
        docname = f'{self.config.project}-{self.config.version}'
        self.write_doc(docname, doctree)
        logger.info('done')

    def write_doc(self, docname: str, doctree: nodes.Element) -> None:
        destination = StringOutput(encoding='utf-8')
        self.writer.write(doctree, destination)
        outfilename = Path(self.outdir) / (os_path(docname) + self.out_suffix)
        ensuredir(str(outfilename.parent))
        try:
            self.writer.save(str(outfilename))
        except OSError as err:
            logger.warning('error writing file %s: %s', outfilename, err)

    def finish(self) -> None:
        pass
