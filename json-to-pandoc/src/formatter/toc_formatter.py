#!/usr/bin/env python3

from helper.logger import *
from helper.latex.latex_section import LatexToCSection

def generate(section_data, section_specs, context, section_index, color_dict, last_section_was_landscape):
    if section_data['section'] != '':
        debug(f"Writing ... {section_data['section'].strip()} {section_data['heading'].strip()}")
    else:
        debug(f"Writing ... {section_data['heading'].strip()}")

    # it is a new section
    section_spec = section_specs.get(section_data.get('section-break'))
    if section_spec is None:
        section_spec = section_specs.get('continuous_portrait')

    latex_section = LatexToCSection(section_data, section_spec, section_index, last_section_was_landscape)
    toc_lines = latex_section.to_latex(color_dict)

    return toc_lines, latex_section.landscape
