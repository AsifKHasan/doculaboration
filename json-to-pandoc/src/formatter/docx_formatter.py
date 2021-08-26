#!/usr/bin/env python3

import textwrap

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
        if 'docx-path' in data['contents']:
            paragraph = doc.add_paragraph('', style='Normal')
            debug('merging docx ; {0}'.format(data['contents']['docx-path']))
            merge_document(paragraph, data['contents']['docx-path'])
