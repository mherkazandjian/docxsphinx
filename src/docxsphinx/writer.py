# -*- coding: utf-8 -*-
"""
    sphinxcontrib-docxwriter
    ~~~~~~~~~~~~~~~~~~~~~~~~~~

    Custom docutils writer for OpenXML (docx).

    :copyright:
        Copyright 2010 by shimizukawa at gmail dot com (Sphinx-users.jp).
    :license: BSD, see LICENSE for details.
"""
import os
import sys
import zipfile
import tempfile
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from docutils import nodes, writers

from sphinx import addnodes
from sphinx.locale import admonitionlabels, versionlabels, _

from docxsphinx import sdocx as docx

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
        try:
            text = dict((k, repr(v)) for k, v in f.f_locals.items()
                        if k != 'self')
            text = unicode(text)
        except:
            text = ''

    if _func is None:
        _func = f.f_code.co_name

    logger.info(' '.join([_func, text]))


class DocxContaner(object):
    pass


class DocxWriter(writers.Writer):
    supported = ('docx',)
    settings_spec = ('No options here.', '', ())
    settings_defaults = {}

    output = None

    def __init__(self, builder):
        writers.Writer.__init__(self)
        self.builder = builder
        self.template_setup()  # setup before call almost docx methods.

        dc = DocxContaner()
        dc.document = docx.newdocument()
        # dc.docbody = dc.document.xpath(
        #         '/w:document/w:body', namespaces=docx.nsprefixes)[0]
        # dc.relationships = docx.relationshiplist()
        # dc.appprops = docx.appproperties()
        # dc.contenttypes = docx.contenttypes()
        # dc.websettings = docx.websettings()
        self.docx_container = dc

    def template_setup(self):
        dotx = self.builder.config['docx_template']
        if dotx:
            # dotx = os.path.join(self.builder.env.srcdir, dotx)
            # z = zipfile.ZipFile(dotx, 'r')
            # template_dir = tempfile.mkdtemp(prefix='docx-')
            # z.extractall(template_dir)
            # docx.set_template(template_dir)
            logger.info("MK using template {}".format(dotx))
            docx.set_template(dotx)

    def save(self, filename):
        dc = self.docx_container
        # wordrelationships = docx.wordrelationships(dc.relationships)
        # coreprops = docx.coreproperties(
        #         title='Python docx demo',
        #         subject='A practical example of making docx from Python',
        #         creator='Mike MacCana',
        #         keywords=['python', 'Office Open XML', 'Word'])
        # docx.savedocx(dc.document, coreprops, dc.appprops, dc.contenttypes,
        #         dc.websettings, wordrelationships, filename)
        dc.document.save(filename)

    def translate(self):
        visitor = DocxTranslator(
                self.document, self.builder, self.docx_container)
        self.document.walkabout(visitor)
        self.output = ''  # visitor.body


class DocxTranslator(nodes.NodeVisitor):

    def __init__(self, document, builder, docx_container):
        self.builder = builder
        self.docx_container = docx_container
        #self.docbody = docx_container.docbody
        self.docbody = docx_container
        nodes.NodeVisitor.__init__(self, document)

        self.states = [[]]
        self.list_style = []
        self.sectionlevel = 0
        self.table = None
        self.list_level = 0
        self.column_widths = None
        self.in_literal_block = False

    def add_text(self, text):
        dprint()
        # HB: cannot print all text in Python 2 because of unicode characters
        # print(text)
        self.states[-1].append(text)

    def new_state(self):
        dprint()
        self.ensure_state()
        self.states.append([])

    def ensure_state(self):
        if self.states and self.states[-1]:
            result = self.states[-1]
            self.states[-1] = []
            # self.docbody.append(
            #         docx.paragraph(''.join(result), breakbefore=True))

            # TODO: Aargh, this needs to be handled properly. OOXML does not have nested lists.
            # So here we close the previous state, which might be a list item. See also depart_list_item
            if self.list_level > 0:
                style = 'List Bullet' if self.list_level < 2 else 'List Bullet {}'.format(self.list_level)
            else:
                style = None

            text = ''.join(result)
            if text:
                self.current_paragraph = self.docbody.document.add_paragraph(
                    text, style=style)
                logger.info('\tensure_state: ')
            # print('\t\t', result)

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

        # HB: I believe this is all wrong, who cares about files?
        #self.docbody.append(docx.pagebreak(type='page', orient='portrait'))
        #self.docbody.document.add_page_break()

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
        self.ensure_state()
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
        #if isinstance(node.parent, nodes.Admonition):
        #    self.add_text(node.astext()+': ')
        #    raise nodes.SkipNode
        self.new_state()

    def depart_title(self, node):
        dprint()
        text = ''.join(self.states.pop())
        dprint(_func='* heading', text=repr(text), level=self.sectionlevel)
        # self.docbody.append(docx.heading(text, self.sectionlevel))
        self.docbody.document.add_heading(text, level=self.sectionlevel)
        # print(text)

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
        self.table[0].append(node['colwidth'])
        #print("HB colwidth {}".format(node['colwidth']))
        # 22, the width of the column in ascii
        raise nodes.SkipNode

    def depart_colspec(self, node):
        dprint()

    def visit_tgroup(self, node):
        dprint()
        pass

    def depart_tgroup(self, node):
        dprint()
        pass

    def visit_thead(self, node):
        dprint()
        pass

    def depart_thead(self, node):
        dprint()
        pass

    def visit_tbody(self, node):
        dprint()
        self.table.append('sep')

    def depart_tbody(self, node):
        dprint()
        pass

    def visit_row(self, node):
        dprint()
        self.table.append([])

    def depart_row(self, node):
        dprint()
        pass

    def visit_entry(self, node):
        dprint()
        if 'morerows' in node or 'morecols' in node:
            raise NotImplementedError('Column or row spanning cells are '
                                      'not implemented.')
        self.new_state()

    def depart_entry(self, node):
        dprint()
        # text = '\n'.join('\n'.join(x) for x in self.states.pop())
        # text = '\n'.join(self.states.pop())
        text = ' '.join(self.states.pop()).strip()
        self.table[-1].append(text)
        #print(text)

    def visit_table(self, node):
        dprint()
        if self.table:
            raise NotImplementedError('Nested tables are not supported.')
        self.new_state()
        #self.current_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        self.table = [[]]

    def depart_table(self, node):
        dprint()
        lines = self.table[1:]
        fmted_rows = []

        # don't allow paragraphs in table cells for now
        for line in lines:
            if line == 'sep':
                pass
            else:
                fmted_rows.append(line)

        nrows = len(fmted_rows)
        ncols = len(fmted_rows[0])
        nncols = [len(r) for r in fmted_rows]

        logger.info("HB {} {}".format(nrows, nncols))
        logger.info("HB {}".format(fmted_rows[0]))

        if self.column_widths is None:
            table = self.docbody.document.add_table(rows=0, cols=ncols, style='Grid Table 4')
        else:
            table = self.docbody.document.add_table(rows=0, cols=0, style='Grid Table 4')
            cols = [table.add_column(Cm(colwidth)) for colwidth in self.column_widths]
            self.column_widths = None

        for row in fmted_rows:
            row_cells = table.add_row().cells
            for i, cell in enumerate(row):
                row_cell = row_cells[i]
                row_cell.text = cell
                para = row_cell.paragraphs[0]
                # TODO: These below should not be here but fixed in the style in the template!
                para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                para.paragraph_format.left_indent = 0


                # row_cells[i].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                #paragraph = row_cells[i].add_paragraph(cell.strip())
                #paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # self.docbody.append(docx.table(fmted_rows))
        self.table = None

        # Add an empty paragraph to prevent tables from being concatenated.
        # TODO: Figure out some better solution.
        self.docbody.document.add_paragraph("")
        self.end_state()

    def visit_acks(self, node):
        dprint()
        raise nodes.SkipNode
        # self.new_state()
        # self.add_text(', '.join(n.astext() for n in node.children[0].children)
        #               + '.')
        # self.end_state()
        raise nodes.SkipNode

    def visit_image(self, node):
        dprint()
        return
        uri = node.attributes['uri']
        file_path = os.path.join(self.builder.env.srcdir, uri)
        dc = self.docx_container
        dc.relationships, picpara = docx.picture(
                dc.relationships, file_path, '')
        self.docbody.append(picpara)

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
        self.new_state()
        self.list_level += 1

    def depart_bullet_list(self, node):
        dprint()
        #self.list_style.pop()
        self.list_level -= 1
        self.end_state()

    def visit_enumerated_list(self, node):
        dprint()
        #self.list_style.append('ListNumber')
        self.new_state()
        self.list_level += 1

    def depart_enumerated_list(self, node):
        dprint()
        #self.list_style.pop()
        self.list_level -= 1
        self.end_state()

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
        # HB: Not sure a new state should be set here, since a new state
        #     is now already set in visit_bullet_list.
        self.new_state()

    def depart_list_item(self, node):
        dprint()
        # HB: should this not use end_state() ?
        text = ''.join(self.states.pop())
        if text:
            ## self.docbody.append(
            ##        docx.paragraph(text, self.list_style[-1], breakbefore=True))
            style = 'List Bullet' if self.list_level < 2 else 'List Bullet {}'.format(self.list_level)
            self.current_paragraph = self.docbody.document.add_paragraph(text, style=style)


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

    def depart_literal_block(self, node):
        dprint()
        #self.end_state()
        text = ''.join(self.states.pop())
        if text:
            ## self.docbody.append(
            ##        docx.paragraph(text, self.list_style[-1], breakbefore=True))
            style = 'Preformatted Text'
            self.current_paragraph = self.docbody.document.add_paragraph(text, style=style)
            self.current_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # self.current_paragraph.paragraph_format.left_indent = 0

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
        self.ensure_state()
        # if not isinstance(node.parent, nodes.Admonition) or \
        #        isinstance(node.parent, addnodes.seealso):
        #     self.new_state()

    def depart_paragraph(self, node):
        dprint()
        # if not isinstance(node.parent, nodes.Admonition) or \
        #        isinstance(node.parent, addnodes.seealso):
        #     self.end_state()

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

    def depart_emphasis(self, node):
        dprint()
        # self.add_text('*')

    def visit_literal_emphasis(self, node):
        dprint()
        # self.add_text('*')

    def depart_literal_emphasis(self, node):
        dprint()
        # self.add_text('*')

    def visit_strong(self, node):
        dprint()
        # self.add_text('**')

    def depart_strong(self, node):
        dprint()
        # self.add_text('**')

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
        #self.add_text(node.astext())
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
