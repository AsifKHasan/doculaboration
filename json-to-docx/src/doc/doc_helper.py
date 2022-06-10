#!/usr/bin/env python3

''' MS Office docx wrapper objects
'''

import time
import yaml
import datetime
import importlib

from docx import Document

from doc.doc_util import *
from helper.logger import *

class DocHelper(object):

    ''' constructor
    '''
    def __init__(self, config):
        self._config = config
        self._doc = Document(self._config['files']['docx-template'])
        self._config['docx'] = self._doc


    ''' generate and save the docx
    '''
    def generate_and_save(self, section_list):
        # we have a concept of nesting_level where parent sections are at level 0 and nested gsheet sections are at subsequent level 1, 2, .....
        nesting_level = 0
        parent_section_index_text = ''

        first_section = True
        section_index = 0

        debug(msg=f"generating docx ..")
        self.start_time = int(round(time.time() * 1000))
        for section in section_list:
            section['nesting-level'] = nesting_level
            section['parent-section-index-text'] = parent_section_index_text

            if section['section'] != '':
                info(f"writing : {section['section'].strip()} {section['heading'].strip()}")
            else:
                info(f"writing : {section['heading'].strip()}")

            section['first-section'] = True if first_section else False
            section['section-index'] = section_index

            module = importlib.import_module("doc.doc_api")
            func = getattr(module, f"process_{section['content-type']}")
            func(section, self._config)

            first_section = False
            section_index = section_index + 1

        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"generating docx .. done {(self.end_time - self.start_time)/1000} seconds")

        # save the docx document
        debug(msg=f"saving docx .. {Path(self._config['files']['output-docx']).resolve()}")
        self.start_time = int(round(time.time() * 1000))
        self._doc.save(self._config['files']['output-docx'])
        set_updatefields_true(self._CONFIG['files']['output-docx'])
        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"saving docx .. done {(self.end_time - self.start_time)/1000} seconds")

        # update indexes
        if sys.platform == 'win32' and self._CONFIG['docx-related']['update-toc']:
            debug(msg=f"updating index .. {Path(self._config['files']['output-docx']).resolve()}")
            self.start_time = int(round(time.time() * 1000))
            update_indexes(self._CONFIG['files']['output-docx'])
            self.end_time = int(round(time.time() * 1000))
            debug(msg=f"updating index .. done {(self.end_time - self.start_time)/1000} seconds")
