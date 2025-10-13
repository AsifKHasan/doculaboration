#!/usr/bin/env python
'''
'''
from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.gsheet.gsheet_util import *

def process(gsheet, section_data, context, current_document_index, nesting_level):
    gsheet_title = section_data['section-prop']['link']
    gsheet_url = section_data['section-prop']['link-target']

    gsheet_id = gsheet_id_from_url(url=gsheet_url, nesting_level=nesting_level)
    info(f"processing gsheet id = [{gsheet_id}] : [{gsheet_title}]", nesting_level=nesting_level)
    
    _gsheethelper = GsheetHelper()
    _data = _gsheethelper.read_gsheet(gsheet_title=gsheet_title, gsheet_url=gsheet_url, parent=section_data, nesting_level=nesting_level+1)

    info(f"processed  gsheet id = [{gsheet_id}] : [{gsheet_title}]", nesting_level=nesting_level)

    return _data
