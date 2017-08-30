# -*- coding: utf-8 -*-
"""
    sphinxcontrib-docxwriter
    ~~~~~~~~~~~~~~~~~~~~~~~~~~

    Custom docutils writer for OpenXML (docx).

    :copyright:
        Copyright 2010 by shimizukawa at gmail dot com (Sphinx-users.jp).
    :license: BSD, see LICENSE for details.
"""
from __future__ import division
import os
import sys

# noinspection PyUnresolvedReferences
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from docx.table import _Cell
from docx import Document

from docutils import nodes, writers

import logging
logging.basicConfig(
    filename='docx.log',
    filemode='w',
    level=logging.INFO,
    format="%(asctime)-15s  %(message)s"
)
logger = logging.getLogger('docx')


def dprint(_func=None, **kw):
    logger.info('-'*50)
    f = sys._getframe(1)
    if kw:
        text = ', '.join('%s = %s' % (k, v) for k, v in kw.items())
    else:
        text = dict((k, repr(v)) for k, v in f.f_locals.items()
                        if k != 'self')
        text = str(text)

    if _func is None:
        _func = f.f_code.co_name

    logger.info(' '.join([_func, text]))


# noinspection PyClassicStyleClass
class DocxWriter(writers.Writer):
    """docutil writer class for docx files"""
    supported = ('docx',)
    settings_spec = ('No options here.', '', ())
    settings_defaults = {}

    output = None
    template_dir = "NO"

    def __init__(self, builder):
        writers.Writer.__init__(self)
        self.builder = builder
        self.template_setup()  # setup before call almost docx methods.

        if self.template_dir == "NO":
            dc = Document()
        else:
            dc = Document(os.path.join('source', self.template_dir))
        self.docx_container = dc

    def template_setup(self):
        dotx = self.builder.config['docx_template']
        if dotx:
            logger.info("MK using template {}".format(dotx))
            self.template_dir = dotx

    def save(self, filename):
        self.docx_container.save(filename)

    def translate(self):
        visitor = DocxTranslator(
                self.document, self.builder, self.docx_container)
        self.document.walkabout(visitor)
        self.output = ''  # visitor.body


# noinspection PyClassicStyleClass
class DocxTranslator(nodes.NodeVisitor):
    """Visitor class to create docx content."""

    def __init__(self, document, builder, docx_container):
        self.builder = builder
        self.docx_container = docx_container
        nodes.NodeVisitor.__init__(self, document)

        self.states = [[]]
        self.list_style = []
        self.sectionlevel = 0
        self.table = None
        self.list_level = 0
        self.column_widths = None
        self.in_literal_block = False
        self.table_style_default = 'Grid Table 4'
        self.table_style = self.table_style_default
        self.more_cols = 0
        self.row = None
        self.cell_counter = 0
        self.strong = False
        self.emphasis = False

        # Self.current_location store places where paragraphs will be added, e.g.
        # typically [document, table-cell]
        self.current_location = [self.docx_container]

        self.ncolumns = 1
        "Number of columns in the current table."

    def add_text(self, text):
        dprint()
        # HB: cannot print all text in Python 2 because of unicode characters
        # print(text)
        textrun = self.current_paragraph.add_run(text)
        if self.strong:
            textrun.bold = True
        if self.emphasis:
            textrun.italic = True

    # TODO: Remove all this state stuff, it should not be necessary anymore
    def new_state(self):
        dprint()
        self.states.append([])

    # TODO: Remove all this state stuff, it should not be necessary anymore
    def end_state(self, first=None):
        dprint()
        result = self.states.pop()
        if first is not None and result:
            item = result[0]
            if item:
                result.insert(0, [first + item[0]])
                result[1] = item[1:]
        self.states[-1].extend(result)
        # print(result)

    def visit_start_of_file(self, node):
        dprint()
        self.new_state()

        # FIXME: visit_start_of_file not close previous section.
        # sectionlevel keep previous and new file's heading level start with
        # previous + 1.
        # This quick hack reset sectionlevel per file.
        # (BTW Sphinx has heading levels per file? or entire document?)
        self.sectionlevel = 0

    def depart_start_of_file(self, node):
        dprint()
        self.end_state()

    def visit_document(self, node):
        dprint()
        self.new_state()

    def depart_document(self, node):
        dprint()
        self.end_state()

    def visit_highlightlang(self, node):
        dprint()
        raise nodes.SkipNode

    def visit_section(self, node):
        dprint()
        self.sectionlevel += 1

    def depart_section(self, node):
        dprint()
        #self.ensure_state()
        if self.sectionlevel > 0:
            self.sectionlevel -= 1

    def visit_topic(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_topic(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    visit_sidebar = visit_topic
    depart_sidebar = depart_topic

    def visit_rubric(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()
        # self.add_text('-[ ')

    def depart_rubric(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text(' ]-')
        # self.end_state()

    def visit_compound(self, node):
        dprint()
        pass

    def depart_compound(self, node):
        dprint()
        pass

    def visit_glossary(self, node):
        dprint()
        pass

    def depart_glossary(self, node):
        dprint()
        pass

    def visit_title(self, node):
        dprint()
        self.new_state()
        self.current_paragraph = self.current_location[-1].add_heading(level=self.sectionlevel)


    def depart_title(self, node):
        dprint()

    def visit_subtitle(self, node):
        dprint()
        pass

    def depart_subtitle(self, node):
        dprint()
        pass

    def visit_attribution(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('-- ')

    def depart_attribution(self, node):
        dprint()
        pass

    def visit_desc(self, node):
        dprint()
        pass

    def depart_desc(self, node):
        dprint()
        pass

    def visit_desc_signature(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()
        # if node.parent['objtype'] in ('class', 'exception'):
        #     self.add_text('%s ' % node.parent['objtype'])

    def depart_desc_signature(self, node):
        dprint()
        raise nodes.SkipNode
        # XXX: wrap signatures in a way that makes sense
        # self.end_state()

    def visit_desc_name(self, node):
        dprint()
        pass

    def depart_desc_name(self, node):
        dprint()
        pass

    def visit_desc_addname(self, node):
        dprint()
        pass

    def depart_desc_addname(self, node):
        dprint()
        pass

    def visit_desc_type(self, node):
        dprint()
        pass

    def depart_desc_type(self, node):
        dprint()
        pass

    def visit_desc_returns(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text(' -> ')

    def depart_desc_returns(self, node):
        dprint()
        pass

    def visit_desc_parameterlist(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('(')
        # self.first_param = 1

    def depart_desc_parameterlist(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text(')')

    def visit_desc_parameter(self, node):
        dprint()
        raise nodes.SkipNode
        # if not self.first_param:
        #     self.add_text(', ')
        # else:
        #     self.first_param = 0
        # self.add_text(node.astext())
        # raise nodes.SkipNode

    def visit_desc_optional(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('[')

    def depart_desc_optional(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text(']')

    def visit_desc_annotation(self, node):
        dprint()
        pass

    def depart_desc_annotation(self, node):
        dprint()
        pass

    def visit_refcount(self, node):
        dprint()
        pass

    def depart_refcount(self, node):
        dprint()
        pass

    def visit_desc_content(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()
        # self.add_text('\n')

    def depart_desc_content(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def visit_figure(self, node):
        # FIXME: figure text become normal paragraph instead of caption.
        dprint()
        self.new_state()

    def depart_figure(self, node):
        dprint()
        self.end_state()

    def visit_caption(self, node):
        dprint()
        pass

    def depart_caption(self, node):
        dprint()
        pass

    def visit_productionlist(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()
        # names = []
        # for production in node:
        #     names.append(production['tokenname'])
        # maxlen = max(len(name) for name in names)
        # for production in node:
        #     if production['tokenname']:
        #         self.add_text(production['tokenname'].ljust(maxlen) + ' ::=')
        #         lastname = production['tokenname']
        #     else:
        #         self.add_text('%s    ' % (' '*len(lastname)))
        #     self.add_text(production.astext() + '\n')
        # self.end_state()
        # raise nodes.SkipNode

    def visit_seealso(self, node):
        dprint()
        self.new_state()

    def depart_seealso(self, node):
        dprint()
        self.end_state(first='')

    def visit_footnote(self, node):
        dprint()
        raise nodes.SkipNode
        # self._footnote = node.children[0].astext().strip()
        # self.new_state()

    def depart_footnote(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state(first='[%s] ' % self._footnote)

    def visit_citation(self, node):
        dprint()
        raise nodes.SkipNode
        # if len(node) and isinstance(node[0], nodes.label):
        #     self._citlabel = node[0].astext()
        # else:
        #     self._citlabel = ''
        # self.new_state()

    def depart_citation(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state(first='[%s] ' % self._citlabel)

    def visit_label(self, node):
        dprint()
        raise nodes.SkipNode

    # XXX: option list could use some better styling

    def visit_option_list(self, node):
        dprint()
        pass

    def depart_option_list(self, node):
        dprint()
        pass

    def visit_option_list_item(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_option_list_item(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def visit_option_group(self, node):
        dprint()
        raise nodes.SkipNode
        # self._firstoption = True

    def depart_option_group(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('     ')

    def visit_option(self, node):
        dprint()
        raise nodes.SkipNode
        # if self._firstoption:
        #     self._firstoption = False
        # else:
        #     self.add_text(', ')

    def depart_option(self, node):
        dprint()
        pass

    def visit_option_string(self, node):
        dprint()
        pass

    def depart_option_string(self, node):
        dprint()
        pass

    def visit_option_argument(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text(node['delimiter'])

    def depart_option_argument(self, node):
        dprint()
        pass

    def visit_description(self, node):
        dprint()
        pass

    def depart_description(self, node):
        dprint()
        pass

    def visit_tabular_col_spec(self, node):
        dprint()
        # TODO: properly implement this!!
        spec = node['spec']
        widths = [float(l.split('cm')[0]) for l in spec.split("{")[1:]]
        self.column_widths = widths
        raise nodes.SkipNode

    def visit_colspec(self, node):
        dprint()
        # The difficulty here is getting the right column width.
        # This can be specified with a tabular_col_spec, see above.
        #
        # Otherwise it is derived from the number of columns, which is
        # defined in visit_tgroup (a bit hackish).
        # The _block_width is the full width of the document, and this
        # is divided by the number of columns.
        #
        # It would perhaps also be possible to use node['colwidth'] in some way.
        #print("HB colwidth {}".format(node['colwidth']))
        # 22, the width of the column in ascii
        if self.column_widths:
            width = self.column_widths[0]
            self.column_widths = self.column_widths[1:]
            col = self.table.add_column(Cm(width))
        else:
            col = self.table.add_column(self.docx_container._block_width // self.ncolumns)

        raise nodes.SkipNode

    def depart_colspec(self, node):
        dprint()

    def visit_tgroup(self, node):
        dprint()
        colspecs = [c for c in node.children if isinstance(c, nodes.colspec)]
        self.ncolumns = len(colspecs)
        print("HB VT {} {} {}".format(self.ncolumns, len(node.children), [type(c) for c in node.children]))
        pass

    def depart_tgroup(self, node):
        dprint()
        self.ncolumns = 1
        pass

    def visit_thead(self, node):
        dprint()
        pass

    def depart_thead(self, node):
        dprint()
        pass

    def visit_tbody(self, node):
        dprint()
        #self.table.append('sep')

    def depart_tbody(self, node):
        dprint()
        pass

    def visit_row(self, node):
        dprint()
        self.row = self.table.add_row()
        self.cell_counter = 0

    def depart_row(self, node):
        dprint()
        pass

    def visit_entry(self, node):
        dprint()
        if 'morerows' in node:
            raise NotImplementedError('Row spanning cells are not implemented.')
        if 'morecols' in node:
            # Hack to make column spanning possible. TODO FIX
            self.more_cols = node['morecols']
        self.new_state()
        cell = self.row.cells[self.cell_counter]
        # A new paragraph will be added by Sphinx, so remove the automated one
        # This turns out to be not possible, so instead the existing one is
        # reused in visit_paragraph.
        #cell.paragraphs.pop()
        if self.more_cols:
            # Perhaps this commented line works no too.
            #cell = cell.merge(self.row.cells[self.cell_counter + self.more_cols])
            for i in range(self.more_cols):
                cell = cell.merge(self.row.cells[self.cell_counter + i + 1])

        self.current_location.append(cell)

    def depart_entry(self, node):
        dprint()
        self.cell_counter = self.cell_counter + self.more_cols + 1
        self.more_cols = 0
        self.current_location.pop()

    def visit_table(self, node):
        dprint()
        # Not sure whether nested tables work now, maybe.
        #if self.table:
        #    raise NotImplementedError('Nested tables are not supported.')
        self.new_state()

        # Columns are added when a colspec is visited.
        try:
            self.table = self.current_location[-1].add_table(
                rows=0, cols=0, style=self.table_style
            )
        except KeyError as exc:
            msg = ('looks like style "{}" is missing\n{}\n'
                   'using no style').format(self.table_style, repr(exc))
            logger.warning(msg)
            self.table = self.current_location[-1].add_table(
                rows=0, cols=0, style=None
            )

    def depart_table(self, node):
        dprint()

        self.table = None
        self.table_style = self.table_style_default

        # Add an empty paragraph to prevent tables from being concatenated.
        # TODO: Figure out some better solution.
        self.docx_container.add_paragraph("")
        self.end_state()

    def visit_acks(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()
        # self.add_text(', '.join(n.astext() for n in node.children[0].children)
        #               + '.')
        # self.end_state()

    def visit_image(self, node):
        dprint()
        uri = node.attributes['uri']
        file_path = os.path.join(self.builder.env.srcdir, uri)
        # TODO implement this.
        return

    def depart_image(self, node):
        dprint()

    def visit_transition(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()
        # self.add_text('=' * 70)
        # self.end_state()

    def visit_bullet_list(self, node):
        dprint()
        # TODO: Apparently it is necessary to take into account whether
        # the list is numbered or not, like the original code did.
        # But that code did not properly account for the level.
        # So merge these two attempts.
        #self.list_style.append('ListBullet')
        #self.new_state()
        self.list_level += 1

    def depart_bullet_list(self, node):
        dprint()
        #self.list_style.pop()
        self.list_level -= 1
        #self.end_state()

    def visit_enumerated_list(self, node):
        dprint()
        #self.list_style.append('ListNumber')
        #self.new_state()
        self.list_level += 1

    def depart_enumerated_list(self, node):
        dprint()
        #self.list_style.pop()
        self.list_level -= 1
        #self.end_state()

    def visit_definition_list(self, node):
        dprint()
        raise nodes.SkipNode
        # self.list_style.append(-2)

    def depart_definition_list(self, node):
        dprint()
        raise nodes.SkipNode
        # self.list_style.pop()

    def visit_list_item(self, node):
        dprint()
        #self.new_state()

        # A new paragraph is created here, but the next visit is to
        # paragraph, so that would add another paragraph. That is
        # prevented if current_paragraph is an empty List paragraph.
        style = 'List Bullet' if self.list_level < 2 else 'List Bullet {}'.format(self.list_level)
        try:
            self.current_paragraph = self.docx_container.add_paragraph(style=style)
        except KeyError as exc:
            msg = ('looks like style "{}" is missing\n{}\n'
                   'using no style').format(style, repr(exc))
            logger.warning(msg)
            self.current_paragraph = self.docx_container.add_paragraph(style=None)

    def depart_list_item(self, node):
        dprint()

    def visit_definition_list_item(self, node):
        dprint()
        raise nodes.SkipNode
        # self._li_has_classifier = len(node) >= 2 and \
        #                           isinstance(node[1], nodes.classifier)

    def depart_definition_list_item(self, node):
        dprint()
        pass

    def visit_term(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_term(self, node):
        dprint()
        raise nodes.SkipNode
        # if not self._li_has_classifier:
        #     self.end_state()

    def visit_classifier(self, node):
        dprint()
        raise nodes.SkipNode
        #self.add_text(' : ')

    def depart_classifier(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def visit_definition(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_definition(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def visit_field_list(self, node):
        dprint()
        pass

    def depart_field_list(self, node):
        dprint()
        pass

    def visit_field(self, node):
        dprint()
        pass

    def depart_field(self, node):
        dprint()
        pass

    def visit_field_name(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_field_name(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text(':')
        # self.end_state()

    def visit_field_body(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_field_body(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def visit_centered(self, node):
        dprint()
        pass

    def depart_centered(self, node):
        dprint()
        pass

    def visit_hlist(self, node):
        dprint()
        pass

    def depart_hlist(self, node):
        dprint()
        pass

    def visit_hlistcol(self, node):
        dprint()
        pass

    def depart_hlistcol(self, node):
        dprint()
        pass

    def visit_admonition(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_admonition(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def _visit_admonition(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def _make_depart_admonition(name):
        def depart_admonition(self, node):
            dprint()
            raise nodes.SkipNode
            # from sphinx.locale import admonitionlabels, versionlabels, _
            # self.end_state(first=admonitionlabels[name] + ': ')
        return depart_admonition

    visit_attention = _visit_admonition
    depart_attention = _make_depart_admonition('attention')
    visit_caution = _visit_admonition
    depart_caution = _make_depart_admonition('caution')
    visit_danger = _visit_admonition
    depart_danger = _make_depart_admonition('danger')
    visit_error = _visit_admonition
    depart_error = _make_depart_admonition('error')
    visit_hint = _visit_admonition
    depart_hint = _make_depart_admonition('hint')
    visit_important = _visit_admonition
    depart_important = _make_depart_admonition('important')
    visit_note = _visit_admonition
    depart_note = _make_depart_admonition('note')
    visit_tip = _visit_admonition
    depart_tip = _make_depart_admonition('tip')
    visit_warning = _visit_admonition
    depart_warning = _make_depart_admonition('warning')

    def visit_versionmodified(self, node):
        dprint()
        raise nodes.SkipNode
        # from sphinx.locale import admonitionlabels, versionlabels, _
        # self.new_state()
        # if node.children:
        #     self.add_text(
        #             versionlabels[node['type']] % node['version'] + ': ')
        # else:
        #     self.add_text(
        #             versionlabels[node['type']] % node['version'] + '.')

    def depart_versionmodified(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def visit_literal_block(self, node):
        dprint()
        # Not sure whether this new_state will work when the literal block
        # is in a list item or a table cell...
        self.new_state()
        self.in_literal_block = True

        # Unlike with Lists, there will not be a visit to paragraph in a
        # literal block, so we *must* create the paragraph here.
        try:
            style = 'Preformatted Text'
            self.current_paragraph = self.docx_container.add_paragraph(style=style)
        except KeyError as exc:
            msg = ('looks like style "{}" is missing\n{}\n'
                   'using no style').format(self.table_style, repr(exc))
            logger.warning(msg)
            style = None
            self.current_paragraph = self.docx_container.add_paragraph(style=style)

        self.current_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def depart_literal_block(self, node):
        dprint()
        self.in_literal_block = False

    def visit_doctest_block(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_doctest_block(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def visit_line_block(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()

    def depart_line_block(self, node):
        dprint()
        raise nodes.SkipNode
        # self.end_state()

    def visit_line(self, node):
        dprint()
        pass

    def depart_line(self, node):
        dprint()
        pass

    def visit_block_quote(self, node):
        # FIXME: working but broken.
        dprint()
        self.new_state()

    def depart_block_quote(self, node):
        dprint()
        self.end_state()

    def visit_compact_paragraph(self, node):
        dprint()
        pass

    def depart_compact_paragraph(self, node):
        dprint()
        pass

    def visit_paragraph(self, node):
        dprint()

        curloc = self.current_location[-1]

        if isinstance(curloc, _Cell):
            if len(curloc.paragraphs):
                if not curloc.paragraphs[0].text:
                    # An empty paragraph is created when a Cell is created.
                    # Reuse this paragraph.
                    self.current_paragraph = curloc.paragraphs[0]
                else:
                    self.current_paragraph = self.current_location[-1].add_paragraph()
            else:
                self.current_paragraph = self.current_location[-1].add_paragraph()
            # HACK because the style is messed up, TODO FIX
            self.current_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            self.current_paragraph.paragraph_format.left_indent = 0
        elif 'List' in self.current_paragraph.style.name and not self.current_paragraph.text:
            # This is the first paragraph in a list item, so do not create another one.
            pass
        else:
            self.current_paragraph = self.current_location[-1].add_paragraph()


    def depart_paragraph(self, node):
        dprint()

    def visit_target(self, node):
        dprint()
        raise nodes.SkipNode

    def visit_index(self, node):
        dprint()
        raise nodes.SkipNode

    def visit_substitution_definition(self, node):
        dprint()
        raise nodes.SkipNode

    def visit_pending_xref(self, node):
        dprint()
        pass

    def depart_pending_xref(self, node):
        dprint()
        pass

    def visit_reference(self, node):
        dprint()
        pass

    def depart_reference(self, node):
        dprint()
        pass

    def visit_download_reference(self, node):
        dprint()
        pass

    def depart_download_reference(self, node):
        dprint()
        pass

    def visit_emphasis(self, node):
        dprint()
        # self.add_text('*')
        self.emphasis = True

    def depart_emphasis(self, node):
        dprint()
        # self.add_text('*')
        self.emphasis = False

    def visit_literal_emphasis(self, node):
        dprint()
        # self.add_text('*')

    def depart_literal_emphasis(self, node):
        dprint()
        # self.add_text('*')

    def visit_strong(self, node):
        dprint()
        # self.add_text('**')
        self.strong = True

    def depart_strong(self, node):
        dprint()
        # self.add_text('**')
        self.strong = False

    def visit_abbreviation(self, node):
        dprint()
        # self.add_text('')

    def depart_abbreviation(self, node):
        dprint()
        # if node.hasattr('explanation'):
        #     self.add_text(' (%s)' % node['explanation'])

    def visit_title_reference(self, node):
        dprint()
        # self.add_text('*')

    def depart_title_reference(self, node):
        dprint()
        # self.add_text('*')

    def visit_literal(self, node):
        dprint()
        # self.add_text('``')

    def depart_literal(self, node):
        dprint()
        # self.add_text('``')

    def visit_subscript(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('_')

    def depart_subscript(self, node):
        dprint()
        pass

    def visit_superscript(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('^')

    def depart_superscript(self, node):
        dprint()
        pass

    def visit_footnote_reference(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('[%s]' % node.astext())

    def visit_citation_reference(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('[%s]' % node.astext())

    def visit_Text(self, node):
        dprint()
        text = node.astext()
        if not self.in_literal_block:
            # assert '\n\n' not in text, 'Found \n\n'
            # Replace double enter with single enter, and single enter with space.
            string_magic = 'TWOENTERSMAGICSTRING'
            text = text.replace('\n\n', string_magic)
            text = text.replace('\n', ' ')
            text = text.replace(string_magic, '\n')
        self.add_text(text)

    def depart_Text(self, node):
        dprint()
        pass

    def visit_generated(self, node):
        dprint()
        pass

    def depart_generated(self, node):
        dprint()
        pass

    def visit_inline(self, node):
        dprint()
        pass

    def depart_inline(self, node):
        dprint()
        pass

    def visit_problematic(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('>>')

    def depart_problematic(self, node):
        dprint()
        raise nodes.SkipNode
        # self.add_text('<<')

    def visit_system_message(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()
        # self.add_text('<SYSTEM MESSAGE: %s>' % node.astext())
        # self.end_state()

    def visit_comment(self, node):
        dprint()
        # TODO: FIX Dirty hack / kludge to set table style.
        # Use proper directives or something like that
        comment = node[0]
        if 'DocxTableStyle' in comment:
            self.table_style = comment.split('DocxTableStyle')[-1].strip()
        print("HB tablestyle {}".format(self.table_style))
        raise nodes.SkipNode

    def visit_meta(self, node):
        dprint()
        raise nodes.SkipNode
        # only valid for HTML

    def visit_raw(self, node):
        dprint()
        raise nodes.SkipNode
        # if 'text' in node.get('format', '').split():
        #     self.body.append(node.astext())

    def unknown_visit(self, node):
        dprint()
        raise nodes.SkipNode
        # raise NotImplementedError('Unknown node: ' + node.__class__.__name__)
