#!/usr/bin/env python3
'''
'''
from helper.logger import *
from helper.gdrive.gdrive_util import *

def process(gsheet, section_data, context, current_document_index):
    warn(f"content type [docx] not supported")
    return None, current_document_index
