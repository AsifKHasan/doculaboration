#!/usr/bin/env python3

import textwrap
import importlib

from helper.logger import *
from helper.pandoc.pandoc_writer import *
from helper.pandoc.pandoc_util import *

def generate(data, doc, section_specs, context):
    use_existing = False
    if data['section-break'] == '-':
        use_existing = True
        data['section-break'] = 'continuous_portrait'

    # TODO: page size and margins
    section = add_section(doc, data, section_specs[data['section-break']], use_existing=use_existing)
 
    if data['section'] != '':
        debug('Writing ... {0} {1}'.format(data['section'], data['heading']).strip())
    else:
        debug('Writing ... {0}'.format(data['heading']).strip())

    return section