#!/usr/bin/env python3

''' Openoffice odt wrapper objects
'''

import time
import yaml
import datetime

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
        self.start_time = int(round(time.time() * 1000))

        # process the sections
        section_list_to_odt(section_list, self._config)

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
