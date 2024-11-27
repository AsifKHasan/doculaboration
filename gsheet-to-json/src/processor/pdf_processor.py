#!/usr/bin/env python3
'''
'''
import pdf2image
import pdf2image.exceptions
from PIL import Image, ImageChops

from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper
from helper.gsheet.gsheet_util import *

def process(gsheet, section_data, context, current_document_index, nesting_level):
    pdf_title = section_data['section-prop']['link']
    pdf_url = section_data['section-prop']['link-target']

    if pdf_url.startswith('https://drive.google.com/file/d/'):
        # the file is from gdrive
        info(f"processing drive file ... [{pdf_title}] : [{pdf_url}]", nesting_level=nesting_level)
        data = download_file_from_drive(pdf_url, context['tmp-dir'], context['drive'], nesting_level=nesting_level+1)

    elif pdf_url.startswith('http'):
        # the file url is a normal web url
        info(f"processing web file ... [{pdf_title}] : [{pdf_url}]", nesting_level=nesting_level)
        data = download_file_from_web(pdf_url, context['tmp-dir'], nesting_level=nesting_level+1)

    else:
        warn(f"the url {pdf_url} is neither a web nor a gdrive url", nesting_level=nesting_level+1)
        data = None

    # extract pages from pdf as images
    if data is not None and 'file-path' in data:
        file_path = data['file-path']
        file_name = data['file-name']
        file_type = data['file-type']
        if file_name.endswith('pdf'):
            file_name = file_name[:-4]

        dpi = 72
        size = None

        # if it is a pdf
        if file_type == 'application/pdf':
            try:
                images = pdf2image.convert_from_path(file_path, fmt='png', dpi=dpi, size=size, transparent=True, output_file=file_name, paths_only=True, output_folder=context['tmp-dir'])
                # mark images for autocropping
                images = [{'path': im_path, 'autocrop': context['autocrop-pdf-pages']} for im_path in images]

            except Exception as e:
                print(e)
                error(f".... could not convert {file_path} to image(s)", nesting_level=nesting_level)

        # if it is an image
        if file_type in ['image/png', 'image/jpeg']:
            images = [{'path': file_path, 'autocrop': False}]

        data['images'] = []
        for image in images:
            im = Image.open(image['path'])
            if image['autocrop']:
                cropped_im = autocrop_image(im, nesting_level=nesting_level+1)
                if cropped_im:
                    im = cropped_im
                    im.save(image['path'])
                    debug(f".... CROPPED image {image['path']}", nesting_level=nesting_level)

                else:
                    debug(f".... image {image['path']} not cropped", nesting_level=nesting_level)

            width, height = im.size
            if 'dpi' in im.info:
                dpi_x, dpi_y = im.info['dpi']
            else:
                dpi_x, dpi_y = 72, 72

            width_in_inches = width / dpi_x
            height_in_inches = height / dpi_y

            data['images'].append({'path': image['path'], 'width': width_in_inches, 'height': height_in_inches})

    return data


''' crop the image automatically 
'''
def autocrop_image(im, nesting_level):
    min_width_height = 2
    debug(f"original image size {im.size}", nesting_level=nesting_level)
    bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
    diff = ImageChops.difference(im, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    # Bounding box given as a 4-tuple defining the left, upper, right, and lower pixel coordinates. If the image is completely empty, this method returns None.
    bbox = diff.getbbox()
    if bbox:
        new_im = im.crop(bbox)
        debug(f"cropped  image size {new_im.size}", nesting_level=nesting_level)
        if new_im.size[0] < min_width_height or new_im.size[1] < min_width_height:
            warn(f"cropped image width/height is less than {(min_width_height)}, will use the original image", nesting_level=nesting_level)
            return None
        else:
            return new_im

    return None
