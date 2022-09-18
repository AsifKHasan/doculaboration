#!/usr/bin/env python3
'''
'''
from collections import defaultdict

import sys
import re
import time
import pygsheets

import urllib.request

from helper.logger import *
from helper.gsheet.gsheet_util import *
from helper.gdrive.gdrive_util import *


def process(gsheet, section_data, context, current_document_index):
    ws_title = section_data['link']

    # if the worksheet has already been read earlier, use the content from cache
    if ws_title in context['worksheet-cache'][gsheet.title]:
        return context['worksheet-cache'][gsheet.title][ws_title]

    info(f"processing ... {gsheet.title} : {ws_title}")
    # get the worksheet data from the context['gsheet-data']
    if ws_title in context['gsheet-data']:
        worksheet_data = context['gsheet-data'][ws_title]
    else:
        warn(f".. {ws_title} not found")
        return {}


    # if any of the cells have userEnteredValue of IMAGE or HYPERLINK, process it
    row = 2
    # start at row 3
    for row_data in worksheet_data['data'][0]['rowData'][2:]:
        val = 1
        if 'values' in row_data:
            # start at column B
            for cell_data in row_data['values'][1:]:
                if 'userEnteredValue' in cell_data:
                    userEnteredValue = cell_data['userEnteredValue']

                    # process where cell contains formulas - image, link to another worksheet, link to another document in gdrive or url
                    if 'formulaValue' in userEnteredValue:
                        formulaValue = userEnteredValue['formulaValue']

                        # content can be an IMAGE/image with an image formula like "=image(....)"
                        m = re.match('=IMAGE\((?P<name>.+)\)', formulaValue, re.IGNORECASE)
                        if m and m.group('name') is not None:
                            row_height = worksheet_data['data'][0]['rowMetadata'][row]['pixelSize']
                            result = download_image(m.group('name'), context['tmp-dir'], row_height)
                            if result:
                                worksheet_data['data'][0]['rowData'][row]['values'][val]['userEnteredValue']['image'] = result

                        # content can be a HYPERLINK/hyperlink to another worksheet
                        m = re.match('=HYPERLINK\("#gid=(?P<ws_gid>.+)",\s*"(?P<ws_title>.+)"\)', formulaValue, re.IGNORECASE)
                        if m and m.group('ws_gid') is not None and m.group('ws_title') is not None:
                            cell_data['contents'] = process(gsheet=gsheet, section_data={'link': m.group('ws_title')}, context=context, current_document_index=current_document_index)

                        # content can be a HYPERLINK/hyperlink to another gdrive file (for now we only allow text only content, that is a text file)
                        m = re.match('=HYPERLINK\("(?P<link_url>.+)",\s*"(?P<link_title>.+)"\)', formulaValue, re.IGNORECASE)
                        if m and m.group('link_url') is not None and m.group('link_title') is not None:
                            url = m.group('link_url')

                            # this may be a drive file
                            if url.startswith('https://drive.google.com/file/d/'):
                                text = read_drive_file(url, context)
                                if text is not None: cell_data['formattedValue'] = text

                            # or it may be a web url
                            elif url.startswith('http'):
                                text = read_web_content(url)
                                if text is not None: cell_data['formattedValue'] = text

                val = val + 1
        row = row + 1

    context['worksheet-cache'][gsheet.title][ws_title] = worksheet_data

    return worksheet_data
