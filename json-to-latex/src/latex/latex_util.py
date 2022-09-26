#!/usr/bin/env python3

'''
various utilities for generating LaTeX code
'''

import sys
import re
import importlib

from helper.logger import *

DEFAULT_FONT = 'Calibri'
# DEFAULT_FONT = 'LiberationSerif'


# font (substitute) map
FONT_MAP = {
    'Calibri': '',
    'Arial': 'Arial',
    'Consolas': 'Consolas',
    'Cambria': 'Cambria',
    'Bree Serif': 'FreeSerif',
}


# LaTeX escape sequences
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


# gsheet border style to LaTeX border style map
GSHEET_LATEX_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dashed',
    'SOLID': 'solid'
}


# gsheet wrap strategy to ConTeXt wrap strategy map
WRAP_STRATEGY_MAP = {'OVERFLOW': 'no-wrap', 'CLIP': 'no-wrap', 'WRAP': 'wrap'}


# gsheet valign to LaTeX taaularray valign mapping
TBLR_VALIGN = {'TOP': 'h', 'MIDDLE': 'm', 'BOTTOM': 'f'}


# gsheet valign to LaTeX paragraph valign mapping
PARA_VALIGN = {'TOP': 't', 'MIDDLE': 'm', 'BOTTOM': 'b'}


# gsheet halign to LaTeX taaularray halign mapping
TBLR_HALIGN = {'LEFT': 'l', 'CENTER': 'c', 'RIGHT': 'r', 'JUSTIFY': 'j'}


# gsheet halign to LaTeX paragraph halign mapping
PARA_HALIGN = {'l': '\raggedright', 'c': '\centering', 'r': '\raggedleft', 'j': 'left'}


# seperation (in inches) between two ConTeXt table columns
COLSEP = (6/72)
# COLSEP = (0/72)


# seperation (in inches) between two ConTeXt table rows
ROWSEP = (2/72)
# ROWSEP = (0/72)


# Horizontal left spacing for first column in header/footer
# HEADER_FOOTER_FIRST_COL_HSPACE = -6
HEADER_FOOTER_FIRST_COL_HSPACE = 0


# Horizontal right spacing for last column in header/footer
# HEADER_FOOTER_LAST_COL_HSPACE = -6
HEADER_FOOTER_LAST_COL_HSPACE = 0


# outline level to ConTeXt style name map
LATEX_HEADING_MAP = {
    'Heading 1' : 'section',
    'Heading 2' : 'subsection',
    'Heading 3' : 'subsubsection',
    'Heading 4' : 'paragraph',
    'Heading 5' : 'subparagraph',
}


# style name to outline level map
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


# outline level to style name map
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


# 0-based gsheet column number to column letter map
COLUMNS = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 
    'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 
    'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 
    'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ', 
]


''' process a list of section_data and generate latex code
'''
def section_list_to_latex(section_list, config, color_dict, document_footnotes):
    section_lines = []
    first_section = True
    for section in section_list:
        section_meta = section['section-meta']
        section_prop = section['section-prop']

        if section_prop['label'] != '':
            info(f"writing : {section_prop['label'].strip()} {section_prop['heading'].strip()}", nesting_level=section_meta['nesting-level'])
        else:
            info(f"writing : {section_prop['heading'].strip()}", nesting_level=section_meta['nesting-level'])

        module = importlib.import_module("latex.latex_api")
        func = getattr(module, f"process_{section_prop['content-type']}")
        section_lines = section_lines + func(section_data=section, config=config, color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' :param path: a path string
    :return: the path that the OS accepts
'''
def os_specific_path(path):
    if sys.platform == 'win32':
        return path.replace('\\', '/')
    else:
        return path



''' fit width/height into a given width/height maintaining aspect ratio
'''
def fit_width_height(fit_within_width, fit_within_height, width_to_fit, height_to_fit):
	WIDTH_OFFSET = 0.0
	HEIGHT_OFFSET = 0.2

	fit_within_width = fit_within_width - WIDTH_OFFSET
	fit_within_height = fit_within_height - HEIGHT_OFFSET

	aspect_ratio = width_to_fit / height_to_fit

	if width_to_fit > fit_within_width:
		width_to_fit = fit_within_width
		height_to_fit = width_to_fit / aspect_ratio
		if height_to_fit > fit_within_height:
			height_to_fit = fit_within_height
			width_to_fit = height_to_fit * aspect_ratio

	return width_to_fit, height_to_fit



''' given pixel size, calculate the row height in inches
    a reasonable approximation is what gsheet says 21 pixels, renders well as 12 pixel (assuming our normal text is 10-11 in size)
'''
def row_height_in_inches(pixel_size):
    return (pixel_size - 9) / 72



''' wrap with BEGIN/END comments
'''
def wrap_with_comment(lines, object_type=None, object_id=None, comment_prefix_start='BEGIN', comment_prefix_stop='END  ', begin_suffix=None, indent_level=0):
    indent = "\t" * indent_level
    output_lines =  list(map(lambda x: f"{indent}{x}", lines))

    if object_type:
        if object_id:
            comment = f"{object_type}: [{object_id}]"

        else:
            comment = f"{object_type}"

        # BEGIN comment
        begin_comment = f"% {comment_prefix_start} {comment}"
        if begin_suffix:
            begin_comment = f"{begin_comment} {begin_suffix}"

        output_lines = [begin_comment] + output_lines

        # END comment
        end_comment = f"% {comment_prefix_stop} {comment}"
        output_lines.append(end_comment)


    return output_lines



''' wrap (in begin/end and indent LaTeX lines
'''
def indent_and_wrap(lines, wrap_in, param_string=None, wrap_prefix_start='begin', wrap_prefix_stop='end', indent_level=1):
    output_lines = []

    # start wrap
    if param_string:
        param_string = param_string.strip('[').strip(']')
        output_lines.append(f"\\{wrap_prefix_start}{wrap_in}[{param_string}]")
    else:
        output_lines.append(f"\\{wrap_prefix_start}{wrap_in}")

    indent = "\t" * indent_level
    output_lines = output_lines + list(map(lambda x: f"{indent}{x}", lines))

    # stop wrap
    output_lines.append(f"\\{wrap_prefix_stop}{wrap_in}")


    return output_lines



''' build LaTeX option string from keywords
'''
def latex_option(*args, **kwargs):
    result = '['

    # iterating over the args list
    for v in args:
        result = result + f"{v}, "

    # iterating over the kwargs dictionary
    for k,v in kwargs.items():
        if v:
            result = result + f"{k}={v}, "

    result = result.strip().strip(',')
    
    result = result + ']'

    return result



''' :param text: a plain text message
    :return: the message escaped to appear correctly in LaTeX
'''
def tex_escape(text):
    regex = re.compile('|'.join(re.escape(str(key)) for key in sorted(CONV.keys(), key = lambda item: - len(item))))
    return regex.sub(lambda match: CONV[match.group()], text)



''' fancy pagestyle
'''
def fancy_pagestyle_header(page_style_name):
    lines = []
    lines.append(f"\\fancypagestyle{{{page_style_name}}}{{")
    lines.append(f"\t\\fancyhf{{}}")
    lines.append(f"\t\\renewcommand{{\\headrulewidth}}{{0pt}}")
    lines.append(f"\t\\renewcommand{{\\footrulewidth}}{{0pt}}")

    return lines



''' insert LaTeX blocks inside text
'''
def process_latex_blocks(text_content):
    # find out if there is any match with LATEX$...$ inside the text_content
    texts_and_latex = []

    pattern = r'LATEX\$[^$]+\$'
    current_index = 0
    for match in re.finditer(pattern, text_content):
        latex_content = match.group()[5:]

        # we have found a LaTeX block, we add the preceding text and the LaTeX block into the list
        latex_start_index, latex_end_index = match.span()[0], match.span()[1]
        if latex_start_index >= current_index:
            # there are preceding text before the latex
            text = text_content[current_index:latex_start_index]

            texts_and_latex.append({'text': text})

            texts_and_latex.append({'latex': latex_content})

            current_index = latex_end_index


    # there may be trailing text
    text = text_content[current_index:]

    texts_and_latex.append({'text': text})

    return texts_and_latex



''' insert footnotes inside text
'''
def process_footnotes(block_id, text_content, document_footnotes, footnote_list):
    # find out if there is any match with FN#key inside the text_content
    # if text contains footnotes we make a list containing texts->footnote->text->footnote ......
    texts_and_footnotes = []

    pattern = r'FN{[^}]+}'
    current_index = 0
    next_footnote_number = len(document_footnotes[block_id]) + 1
    for match in re.finditer(pattern, text_content):
        footnote_key = match.group()[3:-1]
        if footnote_key in footnote_list:
            # debug(f".... footnote {footnote_key} found at {match.span()} with description")
            # we have found a footnote, we add the preceding text and the footnote spec into the list
            footnote_start_index, footnote_end_index = match.span()[0], match.span()[1]
            if footnote_start_index >= current_index:
                # there are preceding text before the footnote
                text = text_content[current_index:footnote_start_index]
                texts_and_footnotes.append({'text': text})

                footnote_mark_latex = f"\\footnotemark[{next_footnote_number}]"
                texts_and_footnotes.append({'fn': footnote_mark_latex})

                # this block has this footnote
                document_footnotes[block_id].append({'key': footnote_key, 'mark': next_footnote_number, 'text': footnote_list[footnote_key]})

                next_footnote_number = next_footnote_number + 1

                current_index = footnote_end_index

        else:
            warn(f".... footnote {footnote_key} found at {match.span()}, but no details found")
            # this is not a footnote, ignore it
            footnote_start_index, footnote_end_index = match.span()[0], match.span()[1]
            # current_index = footnote_end_index + 1

    # there may be trailing text
    text = text_content[current_index:]

    texts_and_footnotes.append({'text': text})

    return texts_and_footnotes



''' process inine contents inside the text
    inine contents are FN{...}, LATEX{...} etc.
'''
def process_inline_blocks(block_id, text_content, document_footnotes, footnote_list, verbatim=False):

    # process FN{...} first, we get a list of block dicts
    inline_blocks = process_footnotes(block_id=block_id, text_content=text_content, document_footnotes=document_footnotes, footnote_list=footnote_list)

    # process LATEX{...} for each text item
    new_inline_blocks = []
    for inline_block in inline_blocks:
        # process only 'text'
        if 'text' in inline_block:
            new_inline_blocks = new_inline_blocks + process_latex_blocks(inline_block['text'])

        else:
            new_inline_blocks.append(inline_block)


    # we are ready prepare the content
    final_content = ''
    for inline_block in new_inline_blocks:
        if 'text' in inline_block:
            if verbatim:
                final_content = final_content + inline_block['text']

            else:
                final_content = final_content + tex_escape(inline_block['text'])

        elif 'fn' in inline_block:
            final_content = final_content + inline_block['fn']

        elif 'latex' in inline_block:
            final_content = final_content + inline_block['latex']


    return final_content