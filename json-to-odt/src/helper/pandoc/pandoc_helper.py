#!/usr/bin/env python3

''' Pandoc wrapper objects
'''

import yaml
import importlib

from helper.pandoc.pandoc_util import *
from helper.logger import *

class Pandoc(object):

    ''' constructor
    '''
    def __init__(self):
        self.document_lines = None


    ''' generate and save the pandoc
    '''
    def generate_pandoc(self, section_list, config, style_path, header_path, pandoc_path):
        # load custom styles for sections, etc.
        sd = yaml.load(open(style_path, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
        section_specs = sd['sections']

        # load header and initialize document
        with open(header_path, "r") as f:
            self.header_lines = [line.rstrip() for line in f]

        self.document_lines = []
        self.color_dict = {}
        section_index = 0
        last_section_was_landscape = True
        for section in section_list:
            content_type = section['content-type']

            # force table formatter for gsheet content
            if content_type == 'gsheet':
                content_type = 'table'

            module = importlib.import_module('formatter.{0}_formatter'.format(content_type))
            section_lines, last_section_was_landscape = module.generate(section, section_specs, config, LETTERS[section_index], self.color_dict, last_section_was_landscape)
            self.document_lines = self.document_lines + section_lines
            section_index = section_index + 1

        # the line before the last line in header_lines is % COLORS, we replace it with set of definecolor's
        for k,v in self.color_dict.items():
            self.header_lines.append(f"\t\definecolor{{{k}}}{{HTML}}{{{v}}}")

        self.header_lines.append("```\n\n")

        # save the markdown document string in a file
        with open(pandoc_path, "w", encoding="utf-8") as f:
            f.write('\n'.join(self.header_lines + self.document_lines))
