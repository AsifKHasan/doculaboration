#!/usr/bin/env python3

'''
various utilities for generating latex code
'''

import sys
import re

from helper.logger import *

DEFAULT_FONT = 'Calibri'
# DEFAULT_FONT = 'LiberationSerif'

FONT_MAP = {
    'Calibri': '',
    'Arial': 'Arial',
    'Consolas': 'Consolas',
    'Cambria': 'Cambria',
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

WRAP_STRATEGY_MAP = {'OVERFLOW': 'no-wrap', 'CLIP': 'no-wrap', 'WRAP': 'wrap'}


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

HEADING_TO_LEVEL = {
    'Heading 1': {'outline-level': 1},
    'Heading 2': {'outline-level': 2},
    'Heading 3': {'outline-level': 3},
    'Heading 4': {'outline-level': 4},
    'Heading 5': {'outline-level': 5},
    'Heading 6': {'outline-level': 6},
    'Heading 7': {'outline-level': 7},
    'Heading 8': {'outline-level': 8},
    'Heading 9': {'outline-level': 9},
    'Heading 10': {'outline-level': 10},
}


LEVEL_TO_HEADING = [
    'Title',
    'Heading 1',
    'Heading 2',
    'Heading 3',
    'Heading 4',
    'Heading 5',
    'Heading 6',
    'Heading 7',
    'Heading 8',
    'Heading 9',
    'Heading 10',
]

LETTERS = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 
    'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 
    'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 
    'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ', 
]


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


''' fancy pagestyle
'''
def fancy_pagestyle_header(page_style_name):
    lines = []
    lines.append(f"\t\\fancypagestyle{{{page_style_name}}}{{")
    lines.append(f"\t\t\\fancyhf{{}}")
    lines.append(f"\t\t\\renewcommand{{\\headrulewidth}}{{0pt}}")
    lines.append(f"\t\t\\renewcommand{{\\footrulewidth}}{{0pt}}")

    return lines


''' insert footnotes inside text
'''
def process_footnotes(text_content, footnote_list, footnote_texts, verbatim=False):
    # find out if there is any match with FN#key inside the text_content
    # if text contains footnotes we make a list containing texts->footnote->text->footnote ......
    texts_and_footnotes = []

    pattern = r'FN{[^}]+}'
    current_index = 0
    for match in re.finditer(pattern, text_content):
        footnote_key = match.group()[3:-1]
        if footnote_key in footnote_list:
            debug(f".... footnote {footnote_key} found at {match.span()} with description")
            # we have found a footnote, we add the preceding text and the footnote spec into the list
            footnote_start_index, footnote_end_index = match.span()[0], match.span()[1]
            if footnote_start_index >= current_index:
                # there are preceding text before the footnote
                text = text_content[current_index:footnote_start_index]
                if not verbatim:
                    text = tex_escape(text)

                texts_and_footnotes.append(text)

                # TODO: footnotemark not supporting any character, it needs an ordered set
                # foot_note_latex = f"\\footnote[{tex_escape(footnote_key)}]{{ {tex_escape(footnote_list[footnote_key])} }}"
                # foot_note_latex = f"\\footnotemark{{{tex_escape(footnote_list[footnote_key])}}}"
                footnote_mark_latex = f"\\footnotemark[{tex_escape(footnote_key)}]"
                texts_and_footnotes.append(footnote_mark_latex)

                footnote_text_latex = f"\\footnotetext[{tex_escape(footnote_key)}]{{{tex_escape(footnote_list[footnote_key])}}}"
                footnote_texts.append(footnote_text_latex)

                current_index = footnote_end_index

        else:
            warn(f".... footnote {footnote_key} found at {match.span()}, but no details found")
            # this is not a footnote, ignore it
            footnote_start_index, footnote_end_index = match.span()[0], match.span()[1]
            # current_index = footnote_end_index + 1

    # there may be trailing text
    text = text_content[current_index:]
    if not verbatim:
        text = tex_escape(text)

    texts_and_footnotes.append(text)

    return ''.join(texts_and_footnotes)
