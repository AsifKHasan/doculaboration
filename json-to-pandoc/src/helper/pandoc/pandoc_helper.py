#!/usr/bin/env python3

'''
Helper to initialize and manipulate pandoc styles, mostly custom styles not present in the templates 
'''

import yaml

from helper.logger import *
from helper.pandoc.pandoc_util import *

HEADER = '''
---
tables: true
geometry:
 - top=20mm
 - left=20mm
...
'''

class PandocHelper(object):

    def __init__(self, style, pandoc_path):
        self._OUTPUT_PANDOC = pandoc_path
        self._STYLE_DATA = style


    def init(self):
        self.load_styles()

        self._doc = HEADER

        return self._doc


    def save(self, doc):
        with open(self._OUTPUT_PANDOC, "w") as f:
            f.write(doc)


    def load_styles(self):
        sd = yaml.load(open(self._STYLE_DATA, 'r', encoding='utf-8'), Loader=yaml.FullLoader)

        self._sections = sd['sections']
