#!/usr/bin/env python3

'''
Helper to initialize and manipulate docbook styles, mostly custom styles not present in the templates 
'''

import yaml

from lxml import etree
from lxml.etree import QName

from helper.logger import *
from helper.docbook.docbook_util import *

class DocbookHelper(object):

    def __init__(self, style, docbook_path):
        self._OUTPUT_DOCBOOK = docbook_path
        self._STYLE_DATA = style


    def init(self):
        self.load_styles()

        xmlns_uris = {None: 'http://docbook.org/ns/docbook'}
        self._doc = etree.Element('article', nsmap=xmlns_uris, version='5.0')
        self._doc.attrib['lang'] = 'en'

        title = etree.SubElement(self._doc, 'title')
        title.text = 'Sample article'
        para = etree.SubElement(self._doc, 'para')
        para.text = 'This is a very short article.'

        return self._doc


    def save(self):
        docbook = etree.ElementTree(self._doc)
        with open(self._OUTPUT_DOCBOOK, "wb") as f:
            docbook.write(f, xml_declaration=True, encoding='utf-8', pretty_print=True)


    def load_styles(self):
        sd = yaml.load(open(self._STYLE_DATA, 'r', encoding='utf-8'), Loader=yaml.FullLoader)

        self._sections = sd['sections']
