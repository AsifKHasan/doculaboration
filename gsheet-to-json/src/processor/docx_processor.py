#!/usr/bin/env python3
'''
'''
from helper.logger import *
from helper.gdrive.gdrive_util import *

def process(link, section_data, context):
    docx_name = section_data['link']
    debug('downloading docx : {0}'.format(link))
    local_path = '{0}/{1}.docx'.format(context['tmp-dir'], docx_name)
    debug('into             : {0}'.format(local_path))
    download_drive_file({'id': link}, local_path, context)
    return {'docx-path': local_path}
