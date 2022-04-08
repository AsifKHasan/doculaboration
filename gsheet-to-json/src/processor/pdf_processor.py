#!/usr/bin/env python3
'''
'''
import pdf2image
import pdf2image.exceptions

from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.gsheet.gsheet_util import *

def process(sheet, section_data, context):
    hyperlink = section_data['link']
    debug(f"processing pdf ... {hyperlink}")

    if hyperlink.startswith('http') and hyperlink.endswith('.pdf'):
        # the pdf url is a normal web url
        data = download_pdf_from_web(hyperlink, context['tmp-dir'])

    elif hyperlink.startswith('https://drive.google.com/file/d/'):
        # the pdf is from gdrive
        data = download_pdf_from_drive(hyperlink, context['tmp-dir'], context['drive'])

    else:
        warn(f"the pdf url {hyperlink} is not either a web or a gdrive url")
        data = None

    # extract pages from pdf as images
    if data is not None and 'pdf_path' in data:
        pdf_file = data['pdf_path']
        pdf_name = data['pdf_name']
        if pdf_name.endswith('pdf'):
            pdf_name = pdf_name[:-4]

        dpi = 150
        size = 1000

        try:
            data['images'] = pdf2image.convert_from_path(pdf_file, fmt='png', dpi=dpi, size=size, transparent=True, output_file=pdf_name, paths_only=True, output_folder=context['tmp-dir'])
        except:
            error(f".... could not convert {pdf_name} to image(s)")

    return data
