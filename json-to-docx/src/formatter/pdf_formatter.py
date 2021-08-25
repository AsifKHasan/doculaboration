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

    # the images go one after another
    if 'contents' in data:
        if data['contents'] and 'images' in data['contents']:
            image_width = page_width
            for image in data['contents']['images']:
                paragraph = doc.add_paragraph('', style='Normal')
                run = paragraph.add_run()
                run.add_picture(image, width=Inches(image_width))
