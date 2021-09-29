#!/usr/bin/env python3

'''
various utilities for generating a pndoc markdown document (mostly latex)
'''

import sys
import lxml
import re
from copy import deepcopy

from helper.logger import *

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


'''
'''
def new_page():
    return wrap_as_latex('\\newpage\n')


''' :param row:
    :return: sum of widths of the merged cells in inches
'''
def merged_cell_span(row, col, start_row, start_col, merges, column_widths):
    cell_width, col_span, row_span = 0, 1, 1
    for merge in merges:
        if merge['startRowIndex'] == (row + start_row) and merge['startColumnIndex'] == (col + start_col):
            col_span = merge['endColumnIndex'] - merge['startColumnIndex']
            row_span = merge['endRowIndex'] - merge['startRowIndex']
            for c in range(col, col + col_span):
                cell_width = cell_width + column_widths[c] + 0.04 * 2

    if cell_width == 0:
        return column_widths[col], col_span, row_span
    else:
        return cell_width, col_span, row_span
