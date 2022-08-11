#!/usr/bin/env python3
'''
'''
import pdf2image
import pdf2image.exceptions
from PIL import Image

from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.gsheet.gsheet_util import *

def process(sheet, section_data, context):
    pdf_title = section_data['link']
    pdf_url = section_data['link-target']
    debug(f"processing file ... {pdf_title} : {pdf_url}")

    if pdf_url.startswith('https://drive.google.com/file/d/'):
        # the file is from gdrive
        data = download_file_from_drive(pdf_url, context['tmp-dir'], context['drive'])

    elif pdf_url.startswith('http'):
        # the file url is a normal web url
        data = download_file_from_web(pdf_url, context['tmp-dir'])

    else:
        warn(f"the url {pdf_url} is neither a web nor a gdrive url")
        data = None

    # extract pages from pdf as images
    if data is not None and 'file-path' in data:
        file_path = data['file-path']
        file_name = data['file-name']
        file_type = data['file-type']
        if file_name.endswith('pdf'):
            file_name = file_name[:-4]

        dpi = 96
        size = None

        # if it is a pdf
        if file_type == 'application/pdf':
            try:
                images = pdf2image.convert_from_path(file_path, fmt='png', dpi=dpi, size=size, transparent=True, output_file=file_name, paths_only=True, output_folder=context['tmp-dir'])
            except Exception as e:
                print(e)
                error(f".... could not convert {pdf_name} to image(s)")

        # if it is an image
        if file_type in ['image/png', 'image/jpeg']:
            images = [file_path]

        data['images'] = []
        for image in images:
            im = Image.open(image)
            width, height = im.size
            if 'dpi' in im.info:
                dpi_x, dpi_y = im.info['dpi']
            else:
                dpi_x, dpi_y = 96, 96

            width_in_inches = width / dpi_x
            height_in_inches = height / dpi_y

            data['images'].append({'path': image, 'width': width_in_inches, 'height': height_in_inches})

    return data
