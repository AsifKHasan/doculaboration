#!/usr/bin/env python

''' MS Office docx wrapper objects
'''

import time
import yaml
import datetime

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
        self.start_time = int(round(time.time() * 1000))

        # override styles
        trace(f"processing custom styles from conf/style-spec.yml")
        self._config['custom-styles'] = {}
        if 'style-specs' in self._config:
            for k, v in self._config['style-specs'].items():
                update_style(doc=self._doc, style_key=k, style_spec=v, custom_styles=self._config['custom-styles'], nesting_level=0)

        # process the sections
        section_list_to_doc(section_list, self._config)

        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"generating docx .. done {(self.end_time - self.start_time)/1000} seconds")

        # save the docx document
        debug(msg=f"saving docx .. {Path(self._config['files']['output-docx']).resolve()}")
        self.start_time = int(round(time.time() * 1000))
        self._doc.save(self._config['files']['output-docx'])
        set_updatefields_true(self._config['files']['output-docx'])
        self.end_time = int(round(time.time() * 1000))
        debug(msg=f"saving docx .. done {(self.end_time - self.start_time)/1000} seconds")

        # update indexes
        if sys.platform == 'win32':
            debug(msg=f"updating index .. {Path(self._config['files']['output-docx']).resolve()}")
            self.start_time = int(round(time.time() * 1000))
            update_indexes(self._config['files']['output-docx'])
            self.end_time = int(round(time.time() * 1000))
            debug(msg=f"updating index .. done {(self.end_time - self.start_time)/1000} seconds")
