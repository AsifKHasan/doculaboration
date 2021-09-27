#!/usr/bin/env python3

import textwrap
import importlib

from helper.logger import *
from helper.pandoc.pandoc_writer import *
from helper.pandoc.pandoc_util import *

from helper.latex.latex_helper import LatexSection

def generate(section_data, section_specs, context):
    if section_data['section'] != '':
        debug('Writing ... {0} {1}'.format(section_data['section'], section_data['heading']).strip())
    else:
        debug('Writing ... {0}'.format(section_data['heading']).strip())

    if section_data['section-break'] == '-':
        section_data['section-break'] = 'continuous_portrait'

    # it is a new section
    latex_section = LatexSection(section_data, section_specs[section_data['section-break']])

    # process contents
    # TODO: for now we just get the latex code for the section, we need to wrap this into a latex document object and append to that document
    section_text = latex_section.to_latex()

    # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
    if 'sections' in section_data['contents']:
        for section in section_data['contents']['sections']:
            content_type = section['content-type']

            # force table formatter for gsheet content
            if content_type == 'gsheet': content_type = 'table'

            module = importlib.import_module('formatter.{0}_formatter'.format(content_type))
            section_text = section_text + module.generate(section, section_specs, context)

    return section_text
