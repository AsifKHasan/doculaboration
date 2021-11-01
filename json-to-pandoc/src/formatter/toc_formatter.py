#!/usr/bin/env python3

from helper.logger import *
from helper.latex.latex_section import LatexToCSection

def generate(section_data, section_specs, context, section_index, color_dict):
    if section_data['section'] != '':
        debug(f"Writing ... {section_data['section'].strip()} {section_data['heading'].strip()}")
    else:
        debug(f"Writing ... {section_data['heading'].strip()}")

    if section_data['section-break'] == '-':
        section_data['section-break'] = 'continuous_portrait'

    # it is a new section
    latex_section = LatexToCSection(section_data, section_specs[section_data['section-break']], section_index)
    toc_lines = latex_section.to_latex(color_dict)

    return toc_lines
