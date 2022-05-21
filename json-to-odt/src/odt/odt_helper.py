#!/usr/bin/env python3

''' Openoffice odt wrapper objects
'''

import time
import yaml
import datetime
import importlib

from odf import opendocument

from odt.odt_util import *
from helper.logger import *

class OdtHelper(object):

    ''' constructor
    '''
    def __init__(self, config):
        self._config = config
        self._odt = opendocument.load(self._config['files']['odt-template'])
        self._config['odt'] = self._odt


    ''' generate and save the odt
    '''
    def generate_and_save(self, section_list):
        # we have a concept of nesting_level where parent sections are at level 0 and nested gsheet sections are at subsequent level 1, 2, .....
        nesting_level = 0
        parent_section_index_text = ''

        first_section = True
        section_index = 0

        debug(msg=f"generating odt ..")
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

            module = importlib.import_module("odt.odt_api")
            func = getattr(module, f"process_{section['content-type']}")
            func(section, self._config)

            first_section = False
            section_index = section_index + 1

        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"generating odt .. done {(self.end_time - self.start_time)/1000} seconds")

        # save the odt document
        debug(msg=f"saving odt .. {Path(self._config['files']['output-odt']).resolve()}")
        self.start_time = int(round(time.time() * 1000))
        self._odt.save(self._config['files']['output-odt'])
        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"saving odt .. done {(self.end_time - self.start_time)/1000} seconds")

        # update indexes
        debug(msg=f"updating index .. {Path(self._config['files']['output-odt']).resolve()}")
        self.start_time = int(round(time.time() * 1000))
        update_indexes(self._odt, self._config['files']['output-odt'])
        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"updating index .. done {(self.end_time - self.start_time)/1000} seconds")
