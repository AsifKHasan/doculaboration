#!/usr/bin/env python
'''
'''

from collections import defaultdict

import sys
import re
import time
import json
import pygsheets

import urllib.request

from helper.logger import *
from helper.util import *


def process(gsheet, section_data, context, current_document_index, nesting_level):
    ws_title = section_data['section-prop']['link']

    # if the worksheet has already been read earlier, use the content from cache
    if ws_title in context['worksheet-cache'][gsheet.title]:
        return context['worksheet-cache'][gsheet.title][ws_title]

    info(f"processing ... [{gsheet.title}] : [{ws_title}]", nesting_level=nesting_level)

    # get the worksheet data from the context['gsheet-data']
    if ws_title in context['gsheet-data'][gsheet.title]:
        worksheet_data = context['gsheet-data'][gsheet.title][ws_title]
    else:
        warn(f"worksheet [{ws_title}] not found", nesting_level=nesting_level)
        return {}

    # if any of the cells have userEnteredValue of IMAGE or HYPERLINK or Range Formula, process it
    row = 2
    # start at row 3
    for row_data in worksheet_data['data'][0]['rowData'][2:]:
        val = 1
        if 'values' in row_data:
            # start at column B
            for cell_data in row_data['values'][1:]:
                # process note
                if 'note' in cell_data:
                    note_json = cell_data['note']
                    tmp_dir = context['tmp-dir']

                    process_note(note_json=note_json, cell_data=cell_data, row=row, val=val, tmp_dir=tmp_dir, context=context, nesting_level=nesting_level)

                if 'userEnteredValue' in cell_data:
                    user_entered_value = cell_data['userEnteredValue']

                    # process where cell contains formulas - image, link to another worksheet, link to another document in gdrive or url
                    if 'formulaValue' in user_entered_value:
                        formula_value = user_entered_value['formulaValue']
                        row_height = worksheet_data['data'][0]['rowMetadata'][row]['pixelSize']
                        tmp_dir = context['tmp-dir']

                        process_formula(formula_value=formula_value, cell_data=cell_data, row=row, val=val, row_height=row_height, tmp_dir=tmp_dir, worksheet_data=worksheet_data, gsheet=gsheet, section_data=section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

                    else:
                        # TODO: 
                        if 'formattedValue' not in cell_data:
                            worksheet_data['data'][0]['rowData'][row]['values'][val]['formattedValue'] = 'xyz'

                val = val + 1

        row = row + 1

    context['worksheet-cache'][gsheet.title][ws_title] = worksheet_data

    return worksheet_data


''' parse formula
'''
def process_note(note_json, cell_data, row, val, tmp_dir, context, nesting_level=0):
    try:
        note_dict = json.loads(note_json)

    except json.JSONDecodeError as e:
        warn(e)
        note_dict = {}
        return

    # include notes into the cell_data for document processing later
    cell_data['notes'] = note_dict

    # check for *background* note and process
    if 'background' in note_dict:
        # NOTE: for now supports image url only
        if 'url' in note_dict['background']:
            url = note_dict['background'].get('url')
            
            # download image
            debug(f"downloading bg image {url}", nesting_level=nesting_level)
            bg_image_dict = download_image(drive_service=context['drive-service'], url=url, tmp_dir=tmp_dir, nesting_level=nesting_level+1)
            cell_data['background'] = bg_image_dict
            debug(f"downloaded  bg image {url}", nesting_level=nesting_level)


''' parse formula
'''
def process_formula(formula_value, cell_data, row, val, row_height, tmp_dir, worksheet_data, gsheet, section_data, context, current_document_index, nesting_level):
    # content can be an IMAGE/image with an image formula like "=image(....)"
    m = re.match(r'=IMAGE\((?P<name>.+)\)', formula_value, re.IGNORECASE)
    if m and m.group('name') is not None:
        result = download_image_from_formula(m.group('name'), tmp_dir, row_height, nesting_level=nesting_level+1)
        if result:
            worksheet_data['data'][0]['rowData'][row]['values'][val]['userEnteredValue']['image'] = result
            cell_data['image'] = result

        return

    # content can be a HYPERLINK/hyperlink to another worksheet
    m = re.match(r'=HYPERLINK\("#gid=(?P<ws_gid>.+)",\s*"(?P<ws_title>.+)"\)', formula_value, re.IGNORECASE)
    if m and m.group('ws_gid') is not None and m.group('ws_title') is not None:
        new_section_data = {'section-prop': {'link': m.group('ws_title')}}
        cell_data['contents'] = process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level+1)

        return

    # content can be a HYPERLINK/hyperlink to another gdrive file (for now we allow pdf and images, but not any kind of text)
    m = re.match(r'=HYPERLINK\("(?P<link_url>.+)",\s*"(?P<link_title>.+)"\)', formula_value, re.IGNORECASE)
    if m and m.group('link_url') is not None and m.group('link_title') is not None:
        url = m.group('link_url')
        title = m.group('link_title')

        # this may be a drive file
        if url.startswith('https://drive.google.com/'):
            info(f"processing drive file ... [{title}] : [{url}]", nesting_level=nesting_level)
            data = download_file_from_drive(drive_service=context['drive-service'], url=url, tmp_dir=context['tmp-dir'], nesting_level=nesting_level+1)
            # if it is an image
            if data['file-type'].startswith('image/'):
                # TODO: we are hardcoding shape, dpi and other parameters for images
                image_params = image_params_from_image(local_path=data['file-path'], row_height=row_height, mode=1, formula=formula_value, formula_parts=[], nesting_level=nesting_level+1)
                if image_params:
                    result =  {**{'url': url, 'path': data['file-path'], 'mode': 1}, **image_params}
                    worksheet_data['data'][0]['rowData'][row]['values'][val]['userEnteredValue']['image'] = result
                    cell_data['image'] = result



        # or it may be a web url
        elif url.startswith('http'):
            text = read_web_content(url, nesting_level=nesting_level+1)
            if text is not None: cell_data['formattedValue'] = text

        return

    # or this can be another range formula which is actually another formula. In this case effectiveValue will be empty
    if 'effectiveValue' in cell_data and cell_data['effectiveValue'] == {}:
        # warn(f"formula [{formula_value}] does not have any corresponsing effectiveValue", nesting_level=nesting_level)

        # check if this is a range formula with the pattern ='worksheet-name'!a1_range
        m = re.match(r"='(?P<ws_name>.+)'!(?P<range>.+)", formula_value, re.IGNORECASE)
        if m and m.group('ws_name') is not None and m.group('range') is not None:
            ws_name = m.group('ws_name')
            range = m.group('range')
            trace(f"formula [{formula_value}] is a link to range [{ws_name}]![{range}]", nesting_level=nesting_level)
            response = get_gsheet_data(google_service=context['service'], gsheet=gsheet, ranges=[f"'{ws_name}'!{range}"])
            range_formula_value = response['sheets'][0]['data'][0]['rowData'][0]['values'][0]['userEnteredValue']['formulaValue']
            process_formula(formula_value=range_formula_value, cell_data=cell_data, row=row, val=val, row_height=row_height, tmp_dir=tmp_dir, worksheet_data=worksheet_data, gsheet=gsheet, section_data=section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

    return
