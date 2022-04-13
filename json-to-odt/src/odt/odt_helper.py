#!/usr/bin/env python3

''' Openoffice odt wrapper objects
'''

import yaml

from odf.opendocument import OpenDocumentText, load
from odf.style import MasterPage, PageLayout, PageLayoutProperties
from odf.text import P

from odt.odt_section import OdtTableSection, OdtToCSection, OdtLoFSection, OdtLoTSection
from odt.odt_util import *
from helper.logger import *

class OdtHelper(object):

    ''' constructor
    '''
    def __init__(self, config):
        self._config = config
        self._odt = load(self._config['files']['odt-template'])
        # self._odt = OpenDocumentText()


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
            section['master-page'] = f"{page_spec}__{margin_spec}__{orientation}"
            master_page = get_or_create_master_page(self._odt, self._config['odt-specs'], section['master-page'], page_spec, margin_spec, orientation)

            # if it is the very first section, change the page-layout of the *Standard* master-page
            if first_section:
                update_standard_master_page(self._odt, section['master-page'])

            this_section_page_spec = self._config['odt-specs']['page-spec'][page_spec]
            this_section_margin_spec = self._config['odt-specs']['margin-spec'][margin_spec]
            section['width'] = float(this_section_page_spec['width']) - float(this_section_margin_spec['left']) - float(this_section_margin_spec['right']) - float(this_section_margin_spec['gutter'])


            first_section = False
            section_index = section_index + 1


    ''' generate and save the odt
    '''
    def generate_and_save(self, section_list):
        self.preprocess(section_list)

        for section in section_list:
            func = getattr(self, f"process_{section['content-type']}")
            func(section)

        # save the odt document
        self._odt.save(self._config['files']['output-odt'])

        # update indexes
        update_indexes(self._odt, self._config['files']['output-odt'])


    ''' Table processor
    '''
    def process_table(self, section_data):
        if section_data['section'] != '':
            debug(f"Writing ... {section_data['section'].strip()} {section_data['heading'].strip()}")
        else:
            debug(f"Writing ... {section_data['heading'].strip()}")

        # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
        if section_data['contents'] is not None and 'sections' in section_data['contents']:
            for section in section_data['contents']['sections']:
                content_type = section['content-type']

                # force table formatter for gsheet content
                if content_type == 'gsheet':
                    content_type = 'table'

                func = getattr(self, f"process_{content_type}")
                func(section)

        else:
            section = OdtTableSection(section_data, self._config)
            section.to_odt(self._odt)


    ''' Table of Content processor
    '''
    def process_toc(self, section_data):
        if section_data['section'] != '':
            debug(f"Writing ... {section_data['section'].strip()} {section_data['heading'].strip()}")
        else:
            debug(f"Writing ... {section_data['heading'].strip()}")

        section = OdtToCSection(section_data, self._config)
        section.to_odt(self._odt)


    ''' List of Figure processor
    '''
    def process_lof(self, section_data):
        if section_data['section'] != '':
            debug(f"Writing ... {section_data['section'].strip()} {section_data['heading'].strip()}")
        else:
            debug(f"Writing ... {section_data['heading'].strip()}")

        section = OdtLoFSection(section_data, self._config)
        section.to_odt(self._odt)


    ''' List of Table processor
    '''
    def process_lot(self, section_data):
        if section_data['section'] != '':
            debug(f"Writing ... {section_data['section'].strip()} {section_data['heading'].strip()}")
        else:
            debug(f"Writing ... {section_data['heading'].strip()}")

        section = OdtLoTSection(section_data, self._config)
        section.to_odt(self._odt)


    ''' pdf processor
    '''
    def process_pdf(self, section_data):
        error(f"content type [pdf] not supported")


    ''' odt processor
    '''
    def process_odt(self, section_data):
        error(f"content type [odt] not supported")


    ''' docx processor
    '''
    def process_docx(self, section_data):
        error(f"content type [docx] not supported")
