#!/usr/bin/env python3

from helper.logger import *
from helper.odt.odt_section import OdtLoFSection

def generate(odt, config, section_data):
    if section_data['section'] != '':
        debug(f"Writing ... {section_data['section'].strip()} {section_data['heading'].strip()}")
    else:
        debug(f"Writing ... {section_data['heading'].strip()}")

    section = OdtLoFSection(section_data, config)
    section.to_odt(odt)
