#!/usr/bin/env python

''' ConTeXt wrapper objects
'''

import time
import yaml
import datetime

from context.context_util import *
from helper.logger import *

class ContextHelper(object):

    ''' constructor
    '''
    def __init__(self, config):
        self._config = config
        self.document_lines = None



    ''' generate and save the ConTeXt
    '''
    def generate_and_save(self, section_list):
        # load header and initialize document
        with open(self._config['files']['document-header'], "r", encoding="utf-8") as f:
            self.header_lines = [line.rstrip() for line in f]

        self.color_dict = {}
        self.headers_footers = {}
        self.document_footnotes = {}
        self.page_layouts = {}

        # process the section-list
        self.document_lines = section_list_to_context(section_list=section_list, config=self._config, color_dict=self.color_dict, headers_footers=self.headers_footers, document_footnotes=self.document_footnotes, page_layouts=self.page_layouts)


        # definecolor
        color_lines = []
        for k,v in self.color_dict.items():
            color_lines.append(f"\\definecolor[{k}][x={v}]")

        # wrap in BEGIN/end comments
        color_lines = wrap_with_comment(lines=color_lines, object_type='Define Colors', indent_level=1)

        color_lines.append("\n")
        self.header_lines = self.header_lines + color_lines


        # Page Layouts
        page_layout_lines = []
        for k, v in self.page_layouts.items():
            page_layout_lines = page_layout_lines + list(map(lambda x: f"\t{x}", v))
            page_layout_lines.append('')

        # wrap in BEGIN/end comments
        page_layout_lines = wrap_with_comment(lines=page_layout_lines, object_type='Page Layouts')

        page_layout_lines.append("\n")
        self.header_lines = self.header_lines + page_layout_lines


        # Header/Footer
        header_footer_lines = []
        for k, v in self.headers_footers.items():
            header_footer_lines = header_footer_lines + list(map(lambda x: f"\t{x}", v))
            header_footer_lines.append('')

        # wrap in BEGIN/end comments
        header_footer_lines = wrap_with_comment(lines=header_footer_lines, object_type='Headers and Footers')

        header_footer_lines.append("\n")
        self.header_lines = self.header_lines + header_footer_lines


        # define the footnote sysmbols through DefineFNsymbols
        # footnote_symbol_lines = []
        # for k, v in self.document_footnotes.items():
        #     if len(v):
        #         footnote_symbol_lines = footnote_symbol_lines + list(map(lambda x: f"\t{x}", define_fn_symbols(name=k, item_list=v)))
        #         footnote_symbol_lines.append('')

        # # wrap in BEGIN/end comments
        # footnote_symbol_lines = wrap_with_comment(lines=footnote_symbol_lines, object_type='Footnote Symbols')

        # footnote_symbol_lines.append("\n")
        # self.header_lines = self.header_lines + footnote_symbol_lines



        # wrap in start/stop text
        self.document_lines = indent_and_wrap(lines=self.document_lines, wrap_in='text', indent_level=1)

        # wrap in BEGIN/end comments
        self.document_lines = wrap_with_comment(lines=self.document_lines, object_type='Text', object_id=None)

        # Document END comment
        self.document_lines.append("% END   Document")


        # save the markdown document string in a file
        with open(self._config['files']['output-context'], "w", encoding="utf-8") as f:
            f.write('\n'.join(self.header_lines + self.document_lines))
