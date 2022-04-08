#!/usr/bin/env python3
'''
'''
from helper.logger import *
from helper.gdrive.gdrive_util import *

def process(link, section_data, context):
    docx_name = section_data['link']
    debug(f"downloading docx : {link}")
    local_path = f"{context['tmp-dir']}/{docx_name}.docx"
    debug(f"into             : {local_path}")
    download_drive_file({'id': link}, local_path, context)
    return {'docx-path': local_path}
