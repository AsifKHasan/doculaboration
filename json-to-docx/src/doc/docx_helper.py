#!/usr/bin/env python

''' MS Office docx wrapper objects
'''

import time
import yaml
import datetime
import pprint

from docx import Document

from helper.config_service import ConfigService
from doc.docx_util import *
from helper.logger import *

class DocxHelper(object):

    ''' constructor
    '''
    def __init__(self, spec_list, nesting_level=0):
        self._docx = Document(ConfigService()._docx_template)

        # specs to config
        ConfigService()._page_specs = spec_list.get('page-spec', {})
        ConfigService()._margin_specs = spec_list.get('margin-spec', {})
        ConfigService()._font_specs = spec_list.get('font-spec', {})
        ConfigService()._style_specs = {}
        for style_key, style_spec in spec_list.get('style-spec', {}).items():
            transfomed_style_spec = transform_nested_dict(data=style_spec, mapping_schema=STYLE_TRANSFORMATION_MAP, nesting_level=nesting_level+1)
            ConfigService()._style_specs[style_key] = transfomed_style_spec

        for k, v in ConfigService()._style_specs.items():
            print(k)
            print('----------')
            print(yaml.dump(v, sort_keys=False))
            print()


    ''' generate and save the docx
    '''
    def generate_and_save(self, section_list, nesting_level=0):
        self.start_time = int(round(time.time() * 1000))

        # override styles
        ConfigService()._custom_styles = {}
        if ConfigService()._style_specs:
            ConfigService()._custom_styles = process_custom_style(docx=self._docx, style_spec=ConfigService()._style_specs, nesting_level=nesting_level+1)

        # process the sections
        section_list_to_docx(docx=self._docx, section_list=section_list, nesting_level=nesting_level+1)

        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"generating docx .. done {(self.end_time - self.start_time)/1000} seconds", nesting_level=nesting_level)

        # save the docx document
        debug(msg=f"saving docx .. {ConfigService()._output_docx_path}", nesting_level=nesting_level)
        self.start_time = int(round(time.time() * 1000))
        self._docx.save(ConfigService()._output_docx_path)
        set_updatefields_true(docx_path=ConfigService()._output_docx_path, nesting_level=nesting_level+1)
        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"saving docx .. done {(self.end_time - self.start_time)/1000} seconds", nesting_level=nesting_level)

        # update indexes
        if sys.platform == 'win32':
            debug(msg=f"updating index .. {ConfigService()._output_docx_path}", nesting_level=nesting_level)
            self.start_time = int(round(time.time() * 1000))
            update_indexes(docx_path=ConfigService()._output_docx_path, nesting_level=nesting_level+1)
            self.end_time = int(round(time.time() * 1000))
            debug(msg=f"updating index .. done {(self.end_time - self.start_time)/1000} seconds", nesting_level=nesting_level)
