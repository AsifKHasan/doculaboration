#!/usr/bin/env python3

''' Openoffice odt wrapper objects
'''

import yaml
import importlib

from odf.opendocument import load

from helper.odt.odt_util import *
from helper.logger import *

class Odt(object):

    ''' constructor
    '''
    def __init__(self, config):
        self._config = config
        self._odt = load(self._config['files']['odt-template'])


    ''' generate and save the odt
    '''
    def generate_odt(self, section_list):
        for section in section_list:
            content_type = section['content-type']

            # force table formatter for gsheet content
            if content_type == 'gsheet':
                content_type = 'table'

            module = importlib.import_module('formatter.{0}_formatter'.format(content_type))
            module.generate(self._odt, self._config, section)

        # save the odt document
        self._odt.save(self._config['files']['output-odt'])
