#!/usr/bin/env python3

import importlib

from helper.logger import *
from helper.latex.latex_section import LatexTableSection

def generate(section_data, section_specs, context, section_index, color_dict, last_section_was_landscape):
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
            section_lines, last_section_was_landscape = module.generate(section, section_specs, context, section_index, color_dict, last_section_was_landscape)

    else:
        # it is a new section
        section_spec = section_specs.get(section_data.get('section-break'))
        if section_spec is None:
            section_spec = section_specs.get('continuous_portrait')

        latex_section = LatexTableSection(section_data, section_spec, section_index, last_section_was_landscape)
        section_lines = latex_section.to_latex(color_dict)
        last_section_was_landscape = latex_section.landscape

    return section_lines, last_section_was_landscape
