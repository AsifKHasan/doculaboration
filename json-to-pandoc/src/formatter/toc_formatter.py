#!/usr/bin/env python3

import textwrap

from helper.logger import *
from helper.pandoc.pandoc_writer import *
from helper.pandoc.pandoc_util import *

def generate(section_data, section_specs, context):
    # it is a new section
    section_text, page_width = add_section(section_data, section_specs[section_data['section-break']])
 
    return section_text
