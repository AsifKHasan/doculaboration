#!/usr/bin/env python3

'''
various utilities for generating a pndoc markdown document (mostly latex)
'''

import sys
import lxml
import re
from copy import deepcopy

from helper.logger import *

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
}

GSHEET_OXML_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dashed',
    'SOLID': 'single',
    'SOLID_MEDIUM': 'thick',
    'SOLID_THICK': 'triple',
    'DOUBLE': 'double',
    'NONE': 'none'
}


''' Table of Contents
'''
def add_toc(doc):
    pass


''' List of Figures
'''
def add_lof(doc):
    pass


''' List of Tables
'''
def add_lot(doc):
    pass


def add_horizontal_line(paragraph, pos='w:bottom', size='6', color='auto'):
    pass

'''
'''
def append_page_number_only(paragraph):
    pass


'''
'''
def append_page_number_with_pages(paragraph, separator=' of '):
    pass


'''
'''
def rotate_text(cell, direction: str):
    pass


'''
'''
def set_character_style(run, spec):
    pass


'''
'''
def cell_bgcolor(color_spec):
    if color_spec:
        red = float(color_spec['red']) if 'red' in color_spec else 0
        green = float(color_spec['green']) if 'green' in color_spec else 0
        blue = float(color_spec['blue']) if 'blue' in color_spec else 0
        s = '\cellcolor[rgb]{{{0},{1},{2}}}'.format(red, green, blue)
    else:
        s = None

    return s

'''
'''
def set_paragraph_bgcolor(paragraph, color):
    pass


'''
'''
def copy_cell_border(from_cell, to_cell):
    pass


'''
'''
def set_paragraph_border(paragraph, **kwargs):
    """
    Set paragraph's border
    Usage:

    set_paragraph_border(
        paragraph,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    pass


'''

'''
def latex_border_from_gsheet_border(borders, side):
    if side in borders:
        border = borders[side]
        red = float(border['color']['red']) if 'red' in border['color'] else 0
        green = float(border['color']['green']) if 'green' in border['color'] else 0
        blue = float(border['color']['blue']) if 'blue' in border['color'] else 0
        width = border['width'] * 0.4
        # TODO: dotted, dashed, double line etc.
        if 'style' in border:
            border_style = border['style']
        else:
            border_style = 'NONE'

        if side in ['left', 'right']:
            return '!{{\\vborder{{{0},{1},{2}}}{{{3}pt}}}}'.format(red, green, blue, width)

        elif side in ['top', 'bottom']:
            return '*1{>{\hborder{1,0,0}{1.0pt}}-}'.format(red, green, blue, width)

            return ''

        else:
            return ''

    else:
        return ''


'''
'''
def insert_image(cell, image_spec):
    '''
        image_spec is like {'url': url, 'path': local_path, 'height': height, 'width': width, 'dpi': im_dpi}
    '''
    if image_spec is not None:
        pass


'''
'''
def set_repeat_table_header(row):
    ''' set repeat table row on every new page
    '''
    return None


'''
'''
def wrap_as_latex(latex):
    return "```{=latex}\n" + latex + "```\n\n"


'''
'''
def start_table(column_sizes):
    str = '''```{=latex}
\setlength\parindent{0pt}

'''

    s = ['p{{{0}in}}'.format(i) for i in column_sizes]
    table_str = '\\begin{{longtable}}[l]{{|{0}|}}\n'.format('|'.join(s))

    return str + table_str + '\n'


'''
'''
def end_table():
    s = '''\end{longtable}
```

'''

    return s


'''
'''
def start_table_row():
    s = '''\TBstrut
'''

    return ''


'''
'''
def end_table_row():
    s = '''\\tabularnewline

'''

    return s


'''
'''
def mark_as_header_row():
    s = '''\endhead
'''
    return s


''' inserts text content into a table cell
    :param text: text to be inserted
    bgcolor: cell bacground color
    left_border:
    right_border:
    top_border:
    bottom_border:
    halign: horizontal alignment of the text inside the cell
    valign: vertical alignment of the text inside the cell
    column_widths: a list of floats describing the column widths of each column in the parent table in inches
    column_number: column index (0 based) of the cell in the parent table
    column_span: how many columns the cell will span to the right including the cell column
    row_span: how many rows the cell will span to the bottom including the cell row
'''
def text_content(text, bgcolor, left_border, right_border, top_border, bottom_border, halign, valign, column_widths, column_number=0, column_span=1, row_span=1):
    # width is the width of all columns under the column span range
    width = sum(column_widths[column_number:column_number + column_span])
    text_latex = '\multicolumn{{{0}}} {{ {1} {2}{{{3}in}} {4} }} {{ {5} {{{6}}} {7} }}\n'.format(column_span, left_border, valign, width, right_border, halign, tex_escape(text), bgcolor)
    text_latex = text_latex if column_number == 0 else '&\n' + text_latex

    top_border_latex = top_border
    bottom_border_latex = bottom_border

    return text_latex, top_border_latex, bottom_border_latex


'''
'''
def image_content(path, halign, valign, column_widths, column_number, column_span, row_span):
    width = sum(column_widths[column_number:column_number + column_span])
    s = '\multicolumn{{{0}}} {{ {1}{{{2}in }}}} {{ {3} {{\includegraphics[width=\\linewidth]{{{4}}}}} }}\n'.format(column_span, valign, width, halign, os_specific_path(path))
    s = s if column_number == 0 else '&\n' + s
    # s = '\includegraphics[width=\\linewidth]{{{0}}}'.format(path)

    return s


'''
'''
def new_page():
    return wrap_as_latex('\\newpage')


'''
'''
def add_section(section_data, section_spec):
    section = ''
    if section_data['section-break'].startswith('newpage_'):
        section = section + '\\newpage' + '\n'

    section = section + '\pdfpagewidth ' + '{0}in'.format(section_spec['page_width']) + '\n'
    section = section + '\pdfpageheight ' + '{0}in'.format(section_spec['page_height']) + '\n'
    section = section + '\\newgeometry{' + 'top={0}in, bottom={0}in, left={0}in, right={0}in'.format(section_spec['top_margin'], section_spec['bottom_margin'], section_spec['left_margin'], section_spec['right_margin']) + '}\n'

    # section.header_distance = Inches(section_spec['header_distance'])
    # section.footer_distance = Inches(section_spec['footer_distance'])
    # section.gutter = Inches(section_spec['gutter'])
    # section.different_first_page_header_footer = section_data['different-first-page-header-footer']

    section = wrap_as_latex(section)



    # get the actual width
    actual_width = float(section_spec['page_width']) - float(section_spec['left_margin']) - float(section_spec['right_margin']) - float(section_spec['gutter'])

    # TODO: set header if it is not set already
    # set_header(doc, section, section_data['header-first'], section_data['header-odd'], section_data['header-even'], actual_width)

    # TODO: set footer if it is not set already
    # set_footer(doc, section, section_data['footer-first'], section_data['footer-odd'], section_data['footer-even'], actual_width)


    if section_data['no-heading'] == False:
        if section_data['section'] != '':
            heading_text = '{0} {1} - {2}'.format('#' * section_data['level'], section_data['section'], section_data['heading']).strip()
        else:
            heading_text = '{0} {1}'.format('#' * section_data['level'], section_data['heading']).strip()

        # headings are styles based on level
        if section_data['level'] != 0:
            section = section + heading_text + '\n\n'
        else:
            s = '\\titlestyle{{{0}}}\n'.format(heading_text)
            s = wrap_as_latex(s)

            section = section + s + '\n'

    return section, actual_width


''' :param text: a plain text message
    :return: the message escaped to appear correctly in LaTeX
'''
def tex_escape(text):
    regex = re.compile('|'.join(re.escape(str(key)) for key in sorted(CONV.keys(), key = lambda item: - len(item))))
    return regex.sub(lambda match: CONV[match.group()], text)


''' :param path: a path string
    :return: the path that the OS accepts
'''
def os_specific_path(path):
    if sys.platform == 'win32':
        return path.replace('\\', '/')
    else:
        return path

''' :param row:
    :return:
'''
def cell_span(row, col, start_row, start_col, merges):
    for merge in merges:
        if merge['startRowIndex'] == (row + start_row) and merge['startColumnIndex'] == (col + start_col):
            col_span = merge['endColumnIndex'] - merge['startColumnIndex']
            row_span = merge['endRowIndex'] - merge['startRowIndex']
            return col_span, row_span

    return 1, 1


''' :param row:
    :return: sum of widths of the merged cells in inches
'''
def merged_cell_width(row, col, start_row, start_col, merges, column_widths):
    cell_width = 0
    for merge in merges:
        if merge['startRowIndex'] == (row + start_row) and merge['startColumnIndex'] == (col + start_col):
            for c in range(col, merge['endColumnIndex'] - start_col):
                cell_width = cell_width + column_widths[c]

    if cell_width == 0:
        return column_widths[col]
    else:
        return cell_width
