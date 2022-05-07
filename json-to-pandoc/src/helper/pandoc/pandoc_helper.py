#!/usr/bin/env python3

''' Pandoc wrapper objects
'''

import yaml

from helper.pandoc.pandoc_util import *
from helper.latex.latex_section import LatexTableSection, LatexToCSection, LatexLoFSection, LatexLoTSection
from helper.logger import *

class Pandoc(object):

    ''' constructor
    '''
    def __init__(self, config):
        self._config = config
        self.document_lines = None


    ''' preprocess document
    '''
    def preprocess(self, section_list):
        first_section = True
        section_index = 0
        for section in section_list:
            section['first-section'] = True if first_section else False
            section['section-index'] = section_index
            section['landscape'] = 'landscape' if section['landscape'] else 'portrait'

            # force table formatter for gsheet content
            if section['content-type'] == 'gsheet':
                section['content-type'] = 'table'

            this_section_page_spec = self._config['odt-specs']['page-spec'][page_spec]
            this_section_margin_spec = self._config['odt-specs']['margin-spec'][margin_spec]
            section['width'] = float(this_section_page_spec['width']) - float(this_section_margin_spec['left']) - float(this_section_margin_spec['right']) - float(this_section_margin_spec['gutter'])

            first_section = False
            section_index = section_index + 1


    ''' generate and save the pandoc
    '''
    def generate_pandoc(self, section_list, config, style_path, header_path, pandoc_path):
        # load custom styles for sections, etc.
        sd = yaml.load(open(style_path, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
        section_specs = sd['sections']

        # load header and initialize document
        with open(header_path, "r") as f:
            self.header_lines = [line.rstrip() for line in f]


        self.preprocess(section_list)

        self.document_lines = []
        self.color_dict = {}
        self.last_section_was_landscape = True
        for section in section_list:
            if section['section'] != '':
                info(f"writing ... {section['section'].strip()} {section['heading'].strip()}")
            else:
                info(f"writing ... {section['heading'].strip()}")

            func = getattr(self, f"process_{section['content-type']}")
            section_lines, self.last_section_was_landscape = func(section)
            self.document_lines = self.document_lines + section_lines

        # the line before the last line in header_lines is % COLORS, we replace it with set of definecolor's
        for k,v in self.color_dict.items():
            self.header_lines.append(f"\t\definecolor{{{k}}}{{HTML}}{{{v}}}")

        self.header_lines.append("```\n\n")

        # save the markdown document string in a file
        with open(pandoc_path, "w", encoding="utf-8") as f:
            f.write('\n'.join(self.header_lines + self.document_lines))


    ''' Table processor
    '''
    def process_table(self, section_data):
        # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
        if section_data['contents'] is not None and 'sections' in section_data['contents']:
            for section in section_data['contents']['sections']:
                content_type = section['content-type']

                # force table formatter for gsheet content
                if content_type == 'gsheet':
                    content_type = 'table'

                func = getattr(self, f"process_{content_type}")
                section_lines, self.last_section_was_landscape = func(section)
                self.document_lines = self.document_lines + section_lines

        else:
            latex_section = LatexTableSection(section_data, self._config, self.last_section_was_landscape)
            section_lines = latex_section.to_latex(self.color_dict)
            self.last_section_was_landscape = latex_section.landscape

        return section_lines, self.last_section_was_landscape
