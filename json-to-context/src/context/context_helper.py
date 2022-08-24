#!/usr/bin/env python3

''' ConTeXt wrapper objects
'''

import time
import yaml
import datetime
import importlib

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
        # we have a concept of nesting_level where parent sections are at level 0 and nested gsheet sections are at subsequent level 1, 2, .....
        nesting_level = 0
        parent_section_index_text = ''

        # load header and initialize document
        with open(self._config['files']['document-header'], "r") as f:
            self.header_lines = [line.rstrip() for line in f]


        first_section = True
        section_index = 0

        self.document_lines = ["% BEGIN Document", "\\begin{document}"]
        self.color_dict = {}
        self.document_footnotes = {}
        for section in section_list:
            section['nesting-level'] = nesting_level
            section['parent-section-index-text'] = parent_section_index_text

            if section['section'] != '':
                info(f"writing : {section['section'].strip()} {section['heading'].strip()}")
            else:
                info(f"writing : {section['heading'].strip()}")

            section['first-section'] = True if first_section else False
            section['section-index'] = section_index

            module = importlib.import_module("context.context_api")
            func = getattr(module, f"process_{section['content-type']}")
            section_lines = func(section_data=section, config=self._config, color_dict=self.color_dict, document_footnotes=self.document_footnotes)
            self.document_lines = self.document_lines + section_lines

            first_section = False
            section_index = section_index + 1
            

        # the line before the last line in header_lines is % COLORS, we replace it with set of definecolor's
        for k,v in self.color_dict.items():
            self.header_lines.append(f"\t\definecolor{{{k}}}{{HTML}}{{{v}}}")

        # define the footnote sysmbols through DefineFNsymbols
        for k,v in self.document_footnotes.items():
            if len(v):
                self.header_lines = self.header_lines + define_fn_symbols(name=k, item_list=v)

        self.header_lines.append("\n")

        
        self.document_lines.append("\n\\end{document}")

        # save the markdown document string in a file
        with open(self._config['files']['output-context'], "w", encoding="utf-8") as f:
            f.write('\n'.join(self.header_lines + self.document_lines))


def define_fn_symbols(name, item_list):
    lines = []
    lines.append(f"")
    lines.append(f"\\DefineFNsymbols{{{name}_symbols}}{{")
    for item in item_list:
        lines.append(f"\t{item['key']}")

    lines.append(f"}}")

    return lines

