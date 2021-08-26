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

    if data['section'] != '':
        debug('Writing ... {0} {1}'.format(data['section'], data['heading']).strip())
    else:
        debug('Writing ... {0}'.format(data['heading']).strip())

    # Heading
    if data['no-heading'] == False:
        paragraph = doc.add_paragraph(data['heading'], style='HEADING1')
        add_horizontal_line(paragraph)

    add_lof(doc)
