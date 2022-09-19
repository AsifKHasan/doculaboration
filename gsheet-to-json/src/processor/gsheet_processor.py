#!/usr/bin/env python3
'''
'''
from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.gsheet.gsheet_util import *

def process(gsheet, section_data, context, current_document_index):
    gsheet_title = section_data['section-prop']['link']
    gsheet_url = section_data['section-prop']['link-target']
    debug(f"processing gsheet ... {gsheet_title} : {gsheet_url}")
    _gsheethelper = GsheetHelper()
    _data, new_document_index = _gsheethelper.read_gsheet(gsheet_title=gsheet_title, gsheet_url=gsheet_url, parent=section_data, current_document_index=current_document_index)
    return _data, new_document_index
