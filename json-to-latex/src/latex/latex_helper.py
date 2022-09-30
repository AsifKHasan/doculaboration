#!/usr/bin/env python3

''' LaTeX wrapper objects
'''

import time
import yaml
import datetime

from latex.latex_util import *
from helper.logger import *

class LatexHelper(object):

    ''' constructor
    '''
    def __init__(self, config):
        self._config = config
        self.document_lines = None



    ''' generate and save the latex
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
        section_lines = section_list_to_latex(section_list=section_list, config=self._config, color_dict=self.color_dict, headers_footers=self.headers_footers, document_footnotes=self.document_footnotes)

        # wrap in BEGIN/end comments
        section_lines = wrap_with_comment(lines=section_lines, object_type='Document', indent_level=1)


        # Page Layouts
        # page_layout_lines = []
        # for k, v in self.page_layouts.items():
        #     page_layout_lines = page_layout_lines + list(map(lambda x: f"\t{x}", v))
        #     page_layout_lines.append('')

        # # wrap in BEGIN/end comments
        # page_layout_lines = wrap_with_comment(lines=page_layout_lines, object_type='Page Layouts')

        # page_layout_lines.append("\n")
        # self.header_lines = self.header_lines + page_layout_lines


        # Header/Footer
        header_footer_lines = []
        for k, v in self.headers_footers.items():
            header_footer_lines = header_footer_lines + v
            header_footer_lines.append('')

        # wrap in BEGIN/end comments
        header_footer_lines = wrap_with_comment(lines=header_footer_lines, object_type='Page Headers and Footers', indent_level=1)

        header_footer_lines.append("\n")
        self.header_lines = self.header_lines + header_footer_lines


        # Colors
        color_lines = []
        for k,v in self.color_dict.items():
            color_lines.append(f"\definecolor{{{k}}}{{HTML}}{{{v}}}")

        # wrap in BEGIN/end comments
        color_lines = wrap_with_comment(lines=color_lines, object_type='Define Colors', indent_level=1)

        color_lines.append("\n")
        self.header_lines = self.header_lines + color_lines


        # define the footnote sysmbols through DefineFNsymbols
        fn_symbol_lines = []
        for k,v in self.document_footnotes.items():
            if len(v):
                fn_symbol_lines = fn_symbol_lines + define_fn_symbols(name=k, item_list=v)

        # wrap in BEGIN/end comments
        fn_symbol_lines = wrap_with_comment(lines=fn_symbol_lines, object_type='Footnote Symbols', indent_level=1)

        fn_symbol_lines.append("\n")
        self.header_lines = self.header_lines + fn_symbol_lines

        
        self.document_lines = ["\\begin{document}"] + section_lines
        self.document_lines.append("\\end{document}")

        # save the markdown document string in a file
        with open(self._config['files']['output-latex'], "w", encoding="utf-8") as f:
            f.write('\n'.join(self.header_lines + self.document_lines))


def define_fn_symbols(name, item_list):
    lines = []
    lines.append(f"")
    lines.append(f"\\DefineFNsymbols{{{name}_symbols}}{{")
    for item in item_list:
        lines.append(f"\t{item['key']}")

    lines.append(f"}}")

    return lines
