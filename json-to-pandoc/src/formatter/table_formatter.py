#!/usr/bin/env python3

import textwrap
import importlib

from helper.logger import *
from helper.pandoc.pandoc_writer import *
from helper.pandoc.pandoc_util import *

def generate(section_data, section_specs, context):
    if section_data['section-break'] == '-':
        section_data['section-break'] = 'continuous_portrait'

    if section_data['section'] != '':
        debug('Writing ... {0} {1}'.format(section_data['section'], section_data['heading']).strip())
    else:
        debug('Writing ... {0}'.format(section_data['heading']).strip())

    # it is a new section
    section_text, page_width = add_section(section_data, section_specs[section_data['section-break']])

    # process contents
    if 'contents' in section_data:
        if section_data['contents']:
            if 'sheets' in section_data['contents']:
                # insert_content (in helper/pandoc/pandoc_writer) is our main work function
                section_text = section_text + insert_content(section_data['contents'], page_width)

            # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
            elif 'sections' in section_data['contents']:
                for section in section_data['contents']['sections']:
                    content_type = section['content-type']

                    # force table formatter for gsheet content
                    if content_type == 'gsheet': content_type = 'table'

                    module = importlib.import_module('formatter.{0}_formatter'.format(content_type))
                    section_text = section_text + module.generate(section, section_specs, context)

    return section_text
