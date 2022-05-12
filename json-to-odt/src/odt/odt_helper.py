#!/usr/bin/env python3

''' Openoffice odt wrapper objects
'''

import yaml

from odf import opendocument, style, text

from odt.odt_section import OdtTableSection, OdtToCSection, OdtLoFSection, OdtLoTSection
from odt.odt_util import *
from helper.logger import *

class OdtHelper(object):

    ''' constructor
    '''
    def __init__(self, config):
        self._config = config
        self._odt = opendocument.load(self._config['files']['odt-template'])


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

            # master-page name
            page_spec = section['page-spec']
            margin_spec = section['margin-spec']
            orientation = section['landscape']
            section['master-page'] = f"mp-{section['section-index']}"
            section['page-layout'] = f"pl-{section['section-index']}"
            master_page = create_master_page(self._odt, self._config['page-specs'], section['master-page'], section['page-layout'], page_spec, margin_spec, orientation)

            # if it is the very first section, change the page-layout of the *Standard* master-page
            if first_section:
                update_master_page_page_layout(self._odt, master_page_name='Standard', new_page_layout_name=section['page-layout'])

            this_section_page_spec = self._config['page-specs']['page-spec'][page_spec]
            this_section_margin_spec = self._config['page-specs']['margin-spec'][margin_spec]
            section['width'] = float(this_section_page_spec['width']) - float(this_section_margin_spec['left']) - float(this_section_margin_spec['right']) - float(this_section_margin_spec['gutter'])

            first_section = False
            section_index = section_index + 1


    ''' generate and save the odt
    '''
    def generate_and_save(self, section_list):
        self.preprocess(section_list)

        for section in section_list:
            if section['section'] != '':
                info(f"writing ... {section['section'].strip()} {section['heading'].strip()}")
            else:
                info(f"writing ... {section['heading'].strip()}")

            func = getattr(self, f"process_{section['content-type']}")
            func(section)

        # save the odt document
        self._odt.save(self._config['files']['output-odt'])

        # update indexes
        update_indexes(self._odt, self._config['files']['output-odt'])


    ''' Table processor
    '''
    def process_table(self, section_data):
        # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
        if section_data['contents'] is not None and 'sections' in section_data['contents']:
            self.preprocess(section_data['contents']['sections'])
            for section in section_data['contents']['sections']:
                content_type = section['content-type']

                # force table formatter for gsheet content
                if content_type == 'gsheet':
                    content_type = 'table'

                debug(f"child content  : index [{section['section-index']}] - [{section['heading']}]")
                func = getattr(self, f"process_{content_type}")
                func(section)

        else:
            debug(f"parent content : index [{section_data['section-index']}] - [{section_data['heading']}]")
            section = OdtTableSection(section_data, self._config)
            section.section_to_odt(self._odt)


    ''' Table of Content processor
    '''
    def process_toc(self, section_data):
        section = OdtToCSection(section_data, self._config)
        section.section_to_odt(self._odt)


    ''' List of Figure processor
    '''
    def process_lof(self, section_data):
        section = OdtLoFSection(section_data, self._config)
        section.section_to_odt(self._odt)


    ''' List of Table processor
    '''
    def process_lot(self, section_data):
        section = OdtLoTSection(section_data, self._config)
        section.section_to_odt(self._odt)


    ''' pdf processor
    '''
    def process_pdf(self, section_data):
        warn(f"content type [pdf] not supported")


    ''' odt processor
    '''
    def process_odt(self, section_data):
        warn(f"content type [odt] not supported")


    ''' docx processor
    '''
    def process_docx(self, section_data):
        warn(f"content type [docx] not supported")
