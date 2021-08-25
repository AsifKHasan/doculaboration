#!/usr/bin/env python3
'''
'''
import pygsheets
from pygsheets.custom_types import *

from helper.logger import *
from helper.gsheet.gsheet_util import *

TABLES = {
    'resource-days': {
        'table-borders': {
            'top': { "style": 'SOLID', "color": hex_to_rgba('808080FF')},
            'right': { "style": 'SOLID', "color": hex_to_rgba('808080FF')},
            'bottom': { "style": 'SOLID', "color": hex_to_rgba('808080FF')},
            'left': { "style": 'SOLID', "color": hex_to_rgba('808080FF')}
        },
        'columns': {
            'A': {'idx': 0, 'size':  60, 'halign': 'CENTER', 'header': '#',              'numberFormat': {'type': 'TEXT'}                               , 'data_index': None},
            'B': {'idx': 1, 'size': 400, 'halign': 'LEFT'  , 'header': 'Roles',          'numberFormat': {'type': 'TEXT'}                               , 'data_index': 0},
            'C': {'idx': 2, 'size': 120, 'halign': 'RIGHT' , 'header': 'Estimated Days', 'numberFormat': {'type': 'CURRENCY', 'pattern': '#,##0.00'}    , 'data_index': 1}
        },
        'header-bg': 'E0E0E0FF',
        'footer-bg': 'F0F0F0FF'
    }
}

def render_resource_days(sheet, ws, current_row, role_days):
    table_def = TABLES['resource-days']
    # header row
    for col_letter, col_data in table_def['columns'].items():
        cell = ws.cell((current_row + 1, col_data['idx'] + 1))
        cell.value = col_data['header']

    # format header
    format_range(sheet, ws, current_row, current_row + 1, 0, 3, bg_color=table_def['header-bg'])

    current_row = current_row + 1

    # data rows
    data_start_row = current_row
    for role_day in role_days:
        for col_letter, col_data in table_def['columns'].items():
            cell = ws.cell((current_row + 1, col_data['idx'] + 1))
            if col_data['data_index'] is not None:
                cell.value = role_day[col_data['data_index']]
            else:
                cell.value = current_row - data_start_row + 1

        current_row = current_row + 1

    data_end_row = current_row - 1

    # footer row, add a formula for Total at col C in row current_row + 1
    current_row = current_row + 1
    cell = ws.cell('C{0}'.format(current_row + 1))
    cell.formula = '=sum(C{0}:C{1})'.format(data_start_row + 1, data_end_row + 1)
    cell = ws.cell('B{0}'.format(current_row + 1))
    cell.value = 'Subtotal'

    # format footer
    format_range(sheet, ws, current_row, current_row + 1, 0, 3, bold=True, bg_color=table_def['footer-bg'])

    # border around whole table
    border_range(sheet, ws, data_start_row - 1, current_row + 1, 0, 3, borders=table_def['table-borders'])

    return current_row

def generate_resource_worksheet(sheet, data, ws_title):
    table_def = TABLES['resource-days']
    # create and init worksheet
    ws = init_worksheet(sheet, ws_title, num_rows=200, num_cols=4, frozen_rows=1, frozen_cols=0, col_def=table_def['columns'])

    # render tables
    current_row = 1
    summary_by_role = data['summary-by-role']
    for k1, v1 in summary_by_role.items():
        # heading 1
        current_row = render_heading1(sheet, ws, current_row, k1)
        # data is list, so start rendering
        if isinstance(v1, list):
            current_row = render_resource_days(sheet, ws, current_row, v1)
            # skip two rows
            current_row = current_row + 2
        # nested dict, data is inside dict
        elif isinstance(v1, dict):
            for k2, v2 in v1.items():
                # heading2
                current_row = render_heading2(sheet, ws, current_row, k2)
                if isinstance(v2, list):
                    current_row = render_resource_days(sheet, ws, current_row, v2)
                    # skip two rows
                    current_row = current_row + 2

def render_heading1(sheet, ws, current_row, heading):
    cell = ws.cell((current_row + 1, 1))
    cell.value = heading
    cell.horizontal_alignment = HorizontalAlignment.LEFT
    cell.wrap_strategy = 'WRAP'
    merge_cells(sheet, ws, current_row, current_row + 1, 0, 3)
    format_range(sheet, ws, current_row, current_row + 1, 0, 3, font_family='Calibri', font_size=12, bold=True, fg_color='6d9eebFF')
    current_row = current_row + 1

    return current_row

def render_heading2(sheet, ws, current_row, heading):
    cell = ws.cell((current_row + 1, 1))
    cell.value = heading
    cell.horizontal_alignment = HorizontalAlignment.LEFT
    cell.wrap_strategy = 'WRAP'
    merge_cells(sheet, ws, current_row, current_row + 1, 0, 3)
    format_range(sheet, ws, current_row, current_row + 1, 0, 3, font_family='Calibri', font_size=11, bold=True, fg_color='980000FF')
    current_row = current_row + 1

    return current_row

def update_sheets(sheets, data):
    sheet = sheets[0]
    generate_resource_worksheet(sheet, data, 'resource-summary')
