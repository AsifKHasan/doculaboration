#!/usr/bin/env python3

'''
Helper to initialize and manipulate docx styles, mostly custom styles not present in the template docx 
'''

import yaml

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

from helper.logger import *
from helper.docx.docx_util import *

class DocxHelper(object):

    def __init__(self, template, style, docx):
        self._STYLE_DATA = style
        self._OUTPUT_DOCX = docx
        self._TEMPLATE_DOCX = template

    def change_builtin_styles(self):
        # change Normal to make it Calibri 10pt
        normal = self._doc.styles['Normal']
        font = normal.font
        font.name = 'Calibri'
        font.size = Pt(10)

        # change Heading 1 so that it starts in a new page always
        h1 = self._doc.styles['Heading 1']
        h1.paragraph_format.page_break_before = True

        # change Heading 2 so that content after it is always together
        h2 = self._doc.styles['Heading 2']
        h2.paragraph_format.keep_with_next = True

    def load_styles(self):
        sd = yaml.load(open(self._STYLE_DATA, 'r', encoding='utf-8'), Loader=yaml.FullLoader)

        for k in sd['paragraph_styles']:
            self.manage_paragraph_style(k, sd['paragraph_styles'][k])

        self._sections = sd['sections']

        self._doc.settings.odd_and_even_pages_header_footer = sd['document']['odd_and_even_pages_header_footer']

    def manage_paragraph_style(self, k, v):
        styles = self._doc.styles
        style = styles.add_style(k, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles[v['base_style']]
        style.next_paragraph_style = styles[v['next_paragraph_style']]

        # character formatting
        font = style.font
        font.name = v['font']['name']
        font.size = Pt(v['font']['size'])
        font.bold = v['font']['bold']
        font.italic = v['font']['italic']
        font.color.rgb = RGBColor.from_string(v['font']['color'])
        font.all_caps = v['font']['all_caps']
        font.small_caps= v['font']['small_caps']

        # paragraph formatting
        pf = style.paragraph_format
        if v['pf']['alignment'] == 'CENTER':
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif v['pf']['alignment'] == 'LEFT':
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif v['pf']['alignment'] == 'RIGHT':
            pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif v['pf']['alignment'] == 'JUSTIFY':
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        pf.first_line_indent = v['pf']['first_line_indent']
        pf.keep_together = v['pf']['keep_together']
        pf.keep_with_next = v['pf']['keep_with_next']
        pf.left_indent = v['pf']['left_indent']
        pf.line_spacing = v['pf']['line_spacing']
        pf.line_spacing_rule = v['pf']['line_spacing_rule']
        pf.page_break_before = v['pf']['page_break_before']
        pf.right_indent = v['pf']['right_indent']
        pf.space_after = Pt(v['pf']['space_after'])
        pf.space_before = Pt(v['pf']['space_before'])
        pf.widow_control = v['pf']['widow_control']

    def init(self):
        self._doc = Document(self._TEMPLATE_DOCX)
        self.change_builtin_styles()
        self.load_styles()
        return self._doc

    def save(self):
        self._doc.save(self._OUTPUT_DOCX)
