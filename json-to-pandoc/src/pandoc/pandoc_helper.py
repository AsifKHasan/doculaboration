#!/usr/bin/env python3

''' Pandoc wrapper objects
'''

import yaml

from pandoc.pandoc_util import *
from latex.latex_section import LatexTableSection, LatexToCSection, LatexLoFSection, LatexLoTSection
from helper.logger import *

class PandocHelper(object):

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

            page_spec = section['page-spec']
            margin_spec = section['margin-spec']

            this_section_page_spec = self._config['page-specs']['page-spec'][page_spec]
            this_section_margin_spec = self._config['page-specs']['margin-spec'][margin_spec]
            section['width'] = float(this_section_page_spec['width']) - float(this_section_margin_spec['left']) - float(this_section_margin_spec['right']) - float(this_section_margin_spec['gutter'])

            first_section = False
            section_index = section_index + 1


    ''' generate and save the pandoc
    '''
    def generate_pandoc(self, section_list):
        # load header and initialize document
        with open(self._config['files']['document-header'], "r") as f:
            self.header_lines = [line.rstrip() for line in f]


        self.preprocess(section_list)

        self.document_lines = []
        self.color_dict = {}
        for section in section_list:
            if section['section'] != '':
                info(f"writing ... {section['section'].strip()} {section['heading'].strip()}")
            else:
                info(f"writing ... {section['heading'].strip()}")

            func = getattr(self, f"process_{section['content-type']}")
            section_lines = func(section)
            self.document_lines = self.document_lines + section_lines

        # the line before the last line in header_lines is % COLORS, we replace it with set of definecolor's
        for k,v in self.color_dict.items():
            self.header_lines.append(f"\t\definecolor{{{k}}}{{HTML}}{{{v}}}")

        self.header_lines.append("```\n\n")

        # save the markdown document string in a file
        with open(self._config['files']['output-pandoc'], "w", encoding="utf-8") as f:
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
                section_lines = func(section)
                self.document_lines = self.document_lines + section_lines

        else:
            latex_section = LatexTableSection(section_data, self._config)
            section_lines = latex_section.to_latex(self.color_dict)

        return section_lines



    ''' Table of Content processor
    '''
    def process_toc(self, section_data):
        latex_section = LatexToCSection(section_data, self._config)
        toc_lines = latex_section.to_latex(self.color_dict)

        return toc_lines

    ''' List of Figure processor
    '''
    def process_lof(self, section_data):
        latex_section = LatexLoFSection(section_data, self._config)
        toc_lines = latex_section.to_latex(self.color_dict)

        return toc_lines


    ''' List of Table processor
    '''
    def process_lot(self, section_data):
        latex_section = LatexLoTSection(section_data, self._config)
        toc_lines = latex_section.to_latex(self.color_dict)

        return toc_lines


    ''' pdf processor
    '''
    def process_pdf(self, section_data):
        warn(f"content type [pdf] not supported")
        return []


    ''' odt processor
    '''
    def process_odt(self, section_data):
        warn(f"content type [odt] not supported")
        return []


    ''' docx processor
    '''
    def process_docx(self, section_data):
        warn(f"content type [docx] not supported")
        return []
