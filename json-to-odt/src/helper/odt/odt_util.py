#!/usr/bin/env python3

''' various utilities for generating an Openoffice odt document
'''

import sys
import re

from odf.style import MasterPage, PageLayout, PageLayoutProperties
from odf.text import P

from helper.logger import *


''' write a paragraph in a given style
'''
def paragraph(odt, style_name, text):
    style = odt.getStyleByName(style_name)
    p = P(stylename=style, text=text)
    odt.text.addElement(p)


''' gets the page-layout from odt if it is there, else create one
'''
def get_or_create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation):
    page_layout = None

    for pl in odt.automaticstyles.getElementsByType(PageLayout):
        if pl.getAttribute('name') == page_layout_name:
            page_layout = pl
            break

    if page_layout is None:
        # create one
        page_layout = PageLayout(name=page_layout_name)
        odt.automaticstyles.addElement(page_layout)
        pageheight = f"{odt_specs['page-spec'][page_spec]['height']}in"
        pagewidth = f"{odt_specs['page-spec'][page_spec]['width']}in"
        margintop = f"{odt_specs['margin-spec'][margin_spec]['top']}in"
        marginbottom = f"{odt_specs['margin-spec'][margin_spec]['bottom']}in"
        marginleft = f"{odt_specs['margin-spec'][margin_spec]['left']}in"
        marginright = f"{odt_specs['margin-spec'][margin_spec]['right']}in"
        # margingutter = f"{odt_specs['margin-spec'][margin_spec]['gutter']}in"
        printorientation = orientation
        page_layout.addElement(PageLayoutProperties(pagewidth=pagewidth, pageheight=pageheight, margintop=margintop, marginbottom=marginbottom, marginleft=marginleft, marginright=marginright, printorientation=orientation))

    return page_layout


''' new odt master-page
    page layouts are saved with a name page-spec__margin-spec__[portrait|landscape]
'''
def get_or_create_master_page(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation):
    # see if the master-page already exists or not in the odt
    master_page = None

    for mp in odt.masterstyles.childNodes:
        if mp.getAttribute('name') == page_layout_name:
            master_page = mp
            break

    if master_page is None:
        # create one, first get/create the page-layout
        page_layout = get_or_create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation)
        master_page = MasterPage(name=page_layout_name, pagelayoutname=page_layout_name)
        odt.masterstyles.addElement(master_page)

    return master_page


LETTERS = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
    'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'
]

DEFAULT_FONT = 'Calibri'

FONT_MAP = {
    'Calibri': '',
    'Arial': 'Arial',
    'Consolas': 'Consolas',
    'Bree Serif': 'FreeSerif',
}

CONV = {
    '&': r'\&',
    '%': r'\%',
    '$': r'\$',
    '#': r'\#',
    '_': r'\_',
    '{': r'\{',
    '}': r'\}',
    '~': r'\textasciitilde{}',
    '^': r'\^{}',
    '\\': r'\textbackslash{}',
    '<': r'\textless{}',
    '>': r'\textgreater{}',
    '\n': r'\linebreak{}',
}

GSHEET_LATEX_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dashed',
    'SOLID': 'solid'
}

TBLR_VALIGN = {'TOP': 'h', 'MIDDLE': 'm', 'BOTTOM': 'f'}
PARA_VALIGN = {'TOP': 't', 'MIDDLE': 'm', 'BOTTOM': 'b'}
TBLR_HALIGN = {'LEFT': 'l', 'CENTER': 'c', 'RIGHT': 'r', 'JUSTIFY': 'j'}
PARA_HALIGN = {'l': '\\raggedright', 'c': '\\centering', 'r': '\\raggedleft'}

COLSEP = (6/72)
ROWSEP = (2/72)

HEADER_FOOTER_FIRST_COL_HSPACE = -6
HEADER_FOOTER_LAST_COL_HSPACE = -6

LATEX_HEADING_MAP = {
    'Heading 1' : 'section',
    'Heading 2' : 'subsection',
    'Heading 3' : 'subsubsection',
    'Heading 4' : 'paragraph',
    'Heading 5' : 'subparagraph',
}


''' :param path: a path string
    :return: the path that the OS accepts
'''
def os_specific_path(path):
    if sys.platform == 'win32':
        return path.replace('\\', '/')
    else:
        return path


''' given pixel size, calculate the row height in inches
    a reasonable approximation is what gsheet says 21 pixels, renders well as 12 pixel (assuming our normal text is 10-11 in size)
'''
def row_height_in_inches(pixel_size):
    return (pixel_size - 9) / 72


''' :param text: a plain text message
    :return: the message escaped to appear correctly in LaTeX
'''
def tex_escape(text):
    regex = re.compile('|'.join(re.escape(str(key)) for key in sorted(CONV.keys(), key = lambda item: - len(item))))
    return regex.sub(lambda match: CONV[match.group()], text)


'''
'''
def mark_as_latex(lines):
    latex_lines = ["```{=latex}"]
    latex_lines = latex_lines + lines
    latex_lines.append("```\n\n")

    return latex_lines
