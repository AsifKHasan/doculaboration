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
        with open(self._config['files']['document-header'], "r") as f:
            self.header_lines = [line.rstrip() for line in f]

        self.document_lines = ["% BEGIN Document", "\\begin{document}"]
        self.color_dict = {}
        self.document_footnotes = {}

        # process the section-list
        self.document_lines = self.document_lines + section_list_to_latex(section_list=section_list, config=self._config, color_dict=self.color_dict, document_footnotes=self.document_footnotes)


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

