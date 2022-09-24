#!/usr/bin/env python3

'''
various utilities for generating ConTeXt code
'''

import sys
import re
import importlib

from helper.logger import *

# distance/gap below header with the content in inches
HEADER_DISTANCE = 0.05

# distance/gap above footer with the content in inches
FOOTER_DISTANCE = 0.05

# distance/gap below footer from the bottom edge of the page
BOTTOM_DISTANCE = 0.00

# distance/gap between right margin and content 
RIGHT_MARGIN_DISTANCE = 0.00

# distance/gap between right margin and edge of page right edge 
RIGHT_EDGE_DISTANCE = 0.00


# height in inches for dummy row
DUMMY_ROW_HEIGHT = 0.10


# default Font
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


# ConTeXt escape sequences
CONV = {
    '\\': r'\backslash',
    '%': r'\%',
    '{': r'\{',
    '}': r'\}',
    '#': r'\#',
    '~': r'\lettertilde',
    '|': r'\|',
    '$': r'\$',
    '_': r'\_',
    '^': r'\letterhat',
    '&': r'\&',
}


# FACTOR by which to multiply gsheet border width to get a reasonable ConTeXt border width (pt)
CONTEXT_BORDER_WIDTH_FACTOR = 0.4

# gsheet border style to ConTeXt border style map
GSHEET_LATEX_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dashed',
    'SOLID': 'solid'
}


# gsheet cell vertical alignment to ConTeXt cell vertical alignment map (bottom high low lohi middle)
CELL_VALIGN_MAP = {
    'TOP': 'high', 
    'MIDDLE': 'lohi', 
    'BOTTOM': 'low'
}


# gsheet cell horizontal alignment to ConTeXt cell horizontal alignment map (flushright flushleft left right center last end)
CELL_HALIGN_MAP = {
    'LEFT': 'flushleft', 
    'CENTER': 'middle', 
    'RIGHT': 'flushright', 
    'JUSTIFY': 'normal'
}


# gsheet image alignment to ConTeXt image alignment map
IMAGE_POSITION = {
    'CENTER': 'middlealigned', 
    'LEFT': 'leftaligned', 
    'RIGHT': 'rightaligned', 
    'TOP': '', 
    'MIDDLE': '', 
    'BOTTOM': ''
}

# gsheet wrap strategy to ConTeXt wrap strategy map
WRAP_STRATEGY_MAP = {
    'OVERFLOW': 'no-wrap', 
    'CLIP': 'no-wrap', 
    'WRAP': 'wrap'
}

# 0-based gsheet column number to column letter map
COLUMNS = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
            'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']


# seperation (in inches) between two ConTeXt table columns
COLSEP = (0/72)

# seperation (in inches) between two ConTeXt table rows
ROWSEP = (0/72)


# outline level to ConTeXt style name map
CONTEXT_HEADING_MAP = {
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

# outline level to ConTeXt style name map
LEVEL_TO_TITLE = [
    'title',
    'chapter',
    'section',
    'subsection',
    'subsubsection',
]


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



''' process a list of section_data and generate context code
'''
def section_list_to_context(section_list, config, color_dict, headers_footers, document_footnotes, page_layouts):
    context_lines = []
    first_section = True
    for section in section_list:
        section_meta = section['section-meta']
        section_prop = section['section-prop']

        if section_prop['label'] != '':
            info(f"writing : {section_prop['label'].strip()} {section_prop['heading'].strip()}", nesting_level=section_meta['nesting-level'])
        else:
            info(f"writing : {section_prop['heading'].strip()}", nesting_level=section_meta['nesting-level'])


        section_meta['first-section'] = first_section
        if first_section:
            first_section = False


        # create the page-layout
        if section_meta['page-layout'] not in page_layouts:
            page_layouts[section_meta['page-layout']] = create_page_layout(page_layout_key=section_meta['page-layout'], page_spec_name=section_prop['page-spec'], landscape=section_prop['landscape'], margin_spec_name=section_prop['margin-spec'], page_specs=config['page-specs'])

        module = importlib.import_module("context.context_api")
        func = getattr(module, f"process_{section_prop['content-type']}")
        context_lines = context_lines + func(section_data=section, config=config, color_dict=color_dict, headers_footers=headers_footers, document_footnotes=document_footnotes, page_layouts=page_layouts)

    return context_lines



''' ConTeXt page_layout from specs
    % A4, portrait, narrow
    \definelayout[a4-portrait-narrow][
        % gutter + left-margin
        backspace=0.50in,
        % page-width - backspace - rightmargin
        width=7.27in,
        % right-margin
        rightmargin=0.50in,

        % top-margin
        topspace=0.25in,
        % header height
        header=0.25in,
        headerdistance=0.10in,
        % page-height - top-margin - bottom-margin
        height=11.19in,
        % footer
        footerdistance=0.10in,
        footer=0.25in,
        % bottom-margin
        bottomdistance=0.00in,
        bottom=0.25in
    ]
'''
def create_page_layout(page_layout_key, page_spec_name, landscape, margin_spec_name, page_specs):
    page_layout_lines = []
    if page_spec_name in page_specs['page-spec'] and margin_spec_name in page_specs['margin-spec']:
        page_spec = page_specs['page-spec'][page_spec_name]
        margin_spec = page_specs['margin-spec'][margin_spec_name]

        if landscape:
            page_layout_lines.append(f"% {page_spec_name}, landscape, {margin_spec_name}")
            width = page_spec['height']
            height = page_spec['width']

        else:
            page_layout_lines.append(f"% {page_spec_name}, portrait, {margin_spec_name}")
            width = page_spec['width']
            height = page_spec['height']

        page_layout_lines.append(f"\\definelayout[{page_layout_key}][")

        left = margin_spec['left']
        right = margin_spec['right']
        top = margin_spec['top']
        bottom = margin_spec['bottom']
        gutter = margin_spec['gutter']

        header_height = margin_spec['header-height']
        footer_height = margin_spec['footer-height']

        page_layout_lines.append(f"\t% gutter + left-margin")
        backspace = gutter + left
        page_layout_lines.append(f"\tbackspace={backspace}in,")

        page_layout_lines.append(f"\t% page-width - backspace - rightmargin")
        content_width = width - backspace - right - RIGHT_MARGIN_DISTANCE - RIGHT_EDGE_DISTANCE
        page_layout_lines.append(f"\twidth={content_width}in,")

        page_layout_lines.append(f"\t% right-margin")
        page_layout_lines.append(f"\trightmargin={right}in,")
        

        page_layout_lines.append(f"\t% top-margin")
        page_layout_lines.append(f"\ttopspace={top}in,")
        
        page_layout_lines.append(f"\t% header height")
        page_layout_lines.append(f"\theader={header_height}in,")
        page_layout_lines.append(f"\theaderdistance={HEADER_DISTANCE}in,")
        
        page_layout_lines.append(f"\t% page-height - top-margin - bottom-margin")
        content_height = height - top - bottom - BOTTOM_DISTANCE
        page_layout_lines.append(f"\theight={content_height}in,")
        
        page_layout_lines.append(f"\t% footer")
        page_layout_lines.append(f"\tfooterdistance={FOOTER_DISTANCE}in,")
        page_layout_lines.append(f"\tfooter={footer_height}in,")
        
        page_layout_lines.append(f"\t% bottom-margin")
        page_layout_lines.append(f"\tbottomdistance={BOTTOM_DISTANCE}in,")
        page_layout_lines.append(f"\tbottom={bottom}in")

        page_layout_lines.append(f"]")


    return page_layout_lines



''' deefine foonote symbols
'''
def define_fn_symbols(name, item_list):
    lines = []
    lines.append(f"\\DefineFNsymbols{{{name}_symbols}}{{")
    for item in item_list:
        lines.append(f"\t{item['key']}")

    lines.append(f"}}")

    return []
    # return lines



''' process line-breaks
'''
def process_line_breaks(text, keep_line_breaks):
    if keep_line_breaks:
        return text.replace('\n', '\\\\ ')

    else:
        return text.replace('\n', ' ')



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



''' build context option string from keywords
'''
def context_option(*args, **kwargs):
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



''' wrap (in start/stop) ad indent ConTeXt lines
'''
def indent_and_wrap(lines, wrap_in, param_string=None, wrap_prefix_start='start', wrap_prefix_stop='stop', indent_level=1):
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



'''
'''
def mark_as_context(lines):
    # context_lines = ["```{=context}"]
    # context_lines = context_lines + lines
    # context_lines.append("```\n\n")

    if len(lines) > 0:
        context_lines = ['']
        context_lines = context_lines + lines
        context_lines.append('')
        return context_lines

    else:
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

        # we have found a ConTeXt block, we add the preceding text and the ConTeXt block into the list
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

                footnote_mark_context = f"\\footnote{{{footnote_list[footnote_key]}}}"
                texts_and_footnotes.append({'fn': footnote_mark_context})

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
def process_inline_blocks(block_id, text_content, document_footnotes, footnote_list, verbatim=False, keep_line_breaks=False):

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
                final_content = process_line_breaks(text=final_content, keep_line_breaks=keep_line_breaks)

        elif 'fn' in inline_block:
            final_content = final_content + inline_block['fn']

        elif 'latex' in inline_block:
            final_content = final_content + inline_block['latex']


    return final_content