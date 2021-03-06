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


def process(sheet, section_data, context):
    ws_title = section_data['link']

    # if the worksheet has already been read earlier, use the content from cache
    if ws_title in context['worksheet-cache'][sheet.title]:
        return context['worksheet-cache'][sheet.title][ws_title]

    info(f"processing ... {sheet.title} : {ws_title}")
    try:
        ws = sheet.worksheet('title', ws_title)
    except:
        warn(f"No worksheet ... {ws_title}")
        return {}

    ranges = [f"{ws_title}!B3:{COLUMNS[ws.cols - 1]}{ws.rows}"]
    include_grid_data = True

    wait_for = context['gsheet-read-wait-seconds']
    try_count = context['gsheet-read-try-count']
    request = context['service'].spreadsheets().get(spreadsheetId=sheet.id, ranges=ranges, includeGridData=include_grid_data)
    response = None
    for i in range(0, try_count):
        try:
            response = request.execute()
            break
        except:
            warn(f"gsheet read request (attempt {i}) failed, waiting for {wait_for} seconds before trying again")
            time.sleep(float(wait_for))

    if response is None:
        error(f"gsheet read request failed, quiting")
        sys.exit(1)

    # if any of the cells have userEnteredValue of IMAGE or HYPERLINK, process it
    row = 0
    for row_data in response['sheets'][0]['data'][0]['rowData']:
        val = 0
        if 'values' in row_data:
            for cell_data in row_data['values']:
                if 'userEnteredValue' in cell_data:
                    userEnteredValue = cell_data['userEnteredValue']

                    # process where cell contains formulas - image, link to another worksheet, link to another document in gdrive or url
                    if 'formulaValue' in userEnteredValue:
                        formulaValue = userEnteredValue['formulaValue']

                        # content can be an IMAGE/image with an image formula like "=image(....)"
                        m = re.match('=IMAGE\((?P<name>.+)\)', formulaValue, re.IGNORECASE)
                        if m and m.group('name') is not None:
                            row_height = response['sheets'][0]['data'][0]['rowMetadata'][row]['pixelSize']
                            result = download_image(m.group('name'), context['tmp-dir'], row_height)
                            if result:
                                response['sheets'][0]['data'][0]['rowData'][row]['values'][val]['userEnteredValue']['image'] = result

                        # content can be a HYPERLINK/hyperlink to another worksheet
                        m = re.match('=HYPERLINK\("#gid=(?P<ws_gid>.+)",\s*"(?P<ws_title>.+)"\)', formulaValue, re.IGNORECASE)
                        if m and m.group('ws_gid') is not None and m.group('ws_title') is not None:
                            # debug(m.group('ws_gid'), m.group('ws_title'))
                            if worksheet_exists(sheet, m.group('ws_title')):
                                cell_data['contents'] = process(sheet, {'link': m.group('ws_title')}, context)

                        # content can be a HYPERLINK/hyperlink to another gdrive file (for now we only allow text only content, that is a text file)
                        m = re.match('=HYPERLINK\("(?P<link_url>.+)",\s*"(?P<link_title>.+)"\)', formulaValue, re.IGNORECASE)
                        if m and m.group('link_url') is not None and m.group('link_title') is not None:
                            url = m.group('link_url')
                            # debug(f".. found a link to [{url}] at [{m.group('link_url')}]")

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

    context['worksheet-cache'][sheet.title][ws_title] = response

    return response
