#!/usr/bin/env python3

import importlib

from helper.logger import *
from helper.odt.odt_section import OdtTableSection

def generate(odt, config, section_data):
    if section_data['section'] != '':
        debug(f"Writing ... {section_data['section'].strip()} {section_data['heading'].strip()}")
    else:
        debug(f"Writing ... {section_data['heading'].strip()}")

    # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
    if section_data['contents'] is not None and 'sections' in section_data['contents']:
        for section in section_data['contents']['sections']:
            content_type = section['content-type']

            # force table formatter for gsheet content
            if content_type == 'gsheet':
                content_type = 'table'

            module = importlib.import_module(f"formatter.{content_type}_formatter")
            module.generate(odt, config, section)

    else:
        section = OdtTableSection(section_data, config)
        section.to_odt(odt)
