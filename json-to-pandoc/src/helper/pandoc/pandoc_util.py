#!/usr/bin/env python3

'''
various utilities for generating a pandoc markdown document
'''

import sys
import re
from copy import deepcopy

from helper.logger import *

''' :param path: a path string
    :return: the path that the OS accepts
'''
def os_specific_path(path):
    if sys.platform == 'win32':
        return path.replace('\\', '/')
    else:
        return path


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
            return '*{{}}{{>{{\\hborder{{{0},{1},{2}}}{{{3}pt}}}}-}}'.format(red, green, blue, width)

        else:
            return ''

    else:
        if side in ['left', 'right']:
            return ''
        elif side in ['top', 'bottom']:
            return '~'
        else:
            return ''


'''
'''
def end_table_row():
    s = '\\tabularnewline\n'

    return s


'''
'''
def mark_as_header_row():
    s = '\endhead\n'

    return s


''' inserts text content into a table cell
    :param text:            text to be inserted
    :param bgcolor:         cell bacground color
    :param left_border:
    :param right_border:
    :param top_border:
    :param bottom_border:
    :param halign:          horizontal alignment of the text inside the cell
    :param valign:          vertical alignment of the text inside the cell
    :param cell_width:      a float describing the cell width in inches
    :param column_number:   column index (0 based) of the cell in the parent table
    :param column_span:     how many columns the cell will span to the right including the cell column
    :param row_span:        how many rows the cell will span to the bottom including the cell row
'''
def text_content(text, bgcolor, left_border, right_border, top_border, bottom_border, halign, valign, cell_width, column_number=0, column_span=1, row_span=1):
    text_latex = '\multicolumn{{{0}}} {{ {1} {2}{{{3}in}} {4} }} {{ {5} {{{6}}} {7} }}\n'.format(column_span, left_border, valign, cell_width, right_border, halign, tex_escape(text), bgcolor)
    text_latex = text_latex if column_number == 0 else '&\n' + text_latex

    top_border_latex = top_border
    bottom_border_latex = bottom_border

    return text_latex, top_border_latex, bottom_border_latex
