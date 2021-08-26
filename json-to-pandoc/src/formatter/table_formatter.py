#!/usr/bin/env python3

import textwrap
import importlib

from helper.logger import *
from helper.docx.docx_writer import *
from helper.docx.docx_util import *

def generate(data, doc, section_specs, context):
    use_existing = False
    if data['section-break'] == '-':
        use_existing = True
        data['section-break'] = 'continuous_portrait'

    section = add_section(doc, data, section_specs[data['section-break']], use_existing=use_existing)
    page_width = section.page_width.inches - section.left_margin.inches - section.right_margin.inches - section.gutter.inches

    if data['section'] != '':
        debug('Writing ... {0} {1}'.format(data['section'], data['heading']).strip())
    else:
        debug('Writing ... {0}'.format(data['heading']).strip())

    if data['no-heading'] == False:
        if data['level'] == 0:
            paragraph = doc.add_paragraph(data['heading'], style='HEADING1')
            add_horizontal_line(paragraph)
        else:
            if data['section'] != '':
                doc.add_heading('{0} - {1}'.format(data['section'], data['heading']).strip(), level=data['level'])
            else:
                doc.add_heading('{0}'.format(data['heading']).strip(), level=data['level'])

    if 'contents' in data:
        if data['contents']:
            if 'sheets' in data['contents']:
                # insert_content (in helper/docx/docx_writer) is our main work function
                insert_content(data['contents'], doc, page_width, None, None)

            # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
            elif 'sections' in data['contents']:
                for section in data['contents']['sections']:
                    content_type = section['content-type']

                    # force table formatter for gsheet content
                    if content_type == 'gsheet': content_type = 'table'

                    module = importlib.import_module('formatter.{0}_formatter'.format(content_type))
                    module.generate(section, doc, section_specs, context)
