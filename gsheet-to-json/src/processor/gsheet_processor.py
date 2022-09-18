#!/usr/bin/env python3
'''
'''
from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.gsheet.gsheet_util import *

def process(gsheet, section_data, context, current_document_index):
    gsheet_title = section_data['link']
    gsheet_url = section_data['link-target']
    debug(f"processing gsheet ... {gsheet_title} : {gsheet_url}")
    _gsheethelper = GsheetHelper()
    new_document_index = current_document_index + 1
    _data = _gsheethelper.read_gsheet(gsheet_title=gsheet_title, gsheet_url=gsheet_url, parent=section_data, current_document_index=new_document_index)
    return _data, new_document_index
