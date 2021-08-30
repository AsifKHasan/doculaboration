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

    if data['section'] != '':
        debug('Writing ... {0} {1}'.format(data['section'], data['heading']).strip())
    else:
        debug('Writing ... {0}'.format(data['heading']).strip())

    # it is a new section
    section = add_section(data, section_specs[data['section-break']])
 
    return section