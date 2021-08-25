#!/usr/bin/env python3
'''
'''
from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.gsheet.gsheet_util import *

def process(sheet, section_data, context):
    gsheet_title = section_data['link']
    debug('processing gsheet ... {0}'.format(gsheet_title))
    _gsheethelper = GsheetHelper()
    _data = _gsheethelper.process_gsheet(gsheet_title, parent=section_data)
    return _data
