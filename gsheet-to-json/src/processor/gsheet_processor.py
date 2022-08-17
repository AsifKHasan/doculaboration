#!/usr/bin/env python3
'''
'''
from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.gsheet.gsheet_util import *

def process(sheet, section_data, context):
    gsheet_title = section_data['link']
    gsheet_url = section_data['link-target']
    debug(f"processing gsheet ... {gsheet_title} : {gsheet_url}")
    _gsheethelper = GsheetHelper()
    _data = _gsheethelper.read_gsheet(gsheet_title=gsheet_title, gsheet_url=gsheet_url, parent=section_data)
    return _data
