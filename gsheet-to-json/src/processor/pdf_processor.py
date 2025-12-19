#!/usr/bin/env python
'''
'''
import cv2
import numpy as np
import pdf2image
import pdf2image.exceptions
from PIL import Image, ImageChops, ImageOps
from pathlib import Path

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

        file_name = file_name + '__'

        # if it is a pdf
        if file_type == 'application/pdf':
            try:
                # split pages into images
                images = pdf2image.convert_from_path(file_path, fmt='jpg', dpi=DPI, size=None, transparent=True, output_file=file_name, paths_only=True, output_folder=context['tmp-dir'])
                
                # mark images for autocropping
                images = [{'path': im_path, 'autocrop': section_data['section-prop']['autocrop']} for im_path in images]

            except Exception as e:
                print(e)
                error(f".... could not convert {file_path} to image(s)", nesting_level=nesting_level)

        # if it is an image
        if file_type in IMAGE_MIME_TYPES:
            images = [{'path': file_path, 'autocrop': section_data['section-prop']['autocrop']}]

        data['images'] = []
        for image in images:
            if image['autocrop']:
                width, height, dpi_x, dpi_y = autocrop_image_pillow(image['path'], nesting_level=nesting_level+1)
                # width, height, dpi_x, dpi_y = autocrop_image_opencv(image['path'], nesting_level=nesting_level+1)
            else:
                width, height, dpi_x, dpi_y = image_meta_pillow(image['path'], nesting_level=nesting_level+1)

            width_in_inches = width / dpi_x
            height_in_inches = height / dpi_y

            data['images'].append({'path': image['path'], 'width': width_in_inches, 'height': height_in_inches})

    return data


''' crop the image automatically
    TODO: looks like it does not work properly
'''
def autocrop_image_pillow(im_path, nesting_level):
    im = Image.open(im_path)
    min_width_height = 2
    width, height = im.size

    if 'dpi' in im.info:
        dpi_x, dpi_y = im.info['dpi']
    else:
        dpi_x, dpi_y = DPI, DPI

    trace(f"original image size [{width}x{height}]", nesting_level=nesting_level)
    bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
    diff = ImageChops.difference(im, bg)
    diff = ImageChops.add(diff, diff, 2.0, 0)
    # Bounding box given as a 4-tuple defining the left, upper, right, and lower pixel coordinates. If the image is completely empty, this method returns None.
    bbox = diff.getbbox()
    if bbox:
        cropped = im.crop(bbox)
        if cropped.size[1] < min_width_height or cropped.size[0] < min_width_height:
            trace(f"cropped image width/height [{cropped.size[0]}x{cropped.size[1]}] is less than {(min_width_height)}, will use the original image", nesting_level=nesting_level)
        else:
            width, height = cropped.size
            trace(f"cropped image: size [{width}x{height}]", nesting_level=nesting_level)
            cropped.save(im_path, dpi=(dpi_x, dpi_y))
            # cropped.save(im_path)

    return width, height, dpi_x, dpi_y


''' crop the image automatically
    TODO: looks like it works better than the pillow version
'''
def autocrop_image_opencv(im_path, nesting_level):
    im = cv2.imread(im_path)
    min_width_height = 2
    height, width, _ = im.shape
    dpi_x, dpi_y = DPI, DPI
    trace(f"original image size [{width}x{height}]", nesting_level=nesting_level)

    gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)

    # Threshold or Canny edge detection to find the subject
    _, thresh = cv2.threshold(gray, 100, 255, cv2.THRESH_BINARY_INV)
    # _, thresh = cv2.threshold(gray, 10, 255, cv2.THRESH_BINARY)
    
    # Find contours
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    if contours:
        # Find the largest contour by area
        # c = max(contours, key=cv2.contourArea)
        # x, y, w, h = cv2.boundingRect(c)

        # Combine all contours into one bounding box
        x, y, w, h = cv2.boundingRect(np.vstack(contours))
        
        # Crop the image based on the bounding box
        if h < min_width_height or w < min_width_height:
            trace(f"cropped image width/height [{w}x{h}] is less than {(min_width_height)}, will use the original image", nesting_level=nesting_level)
        else:
            cropped = im[y:y+h, x:x+w]
            height, width, _ = cropped.shape
            cv2.imwrite(im_path, cropped)
            trace(f"cropped image: size [{w}x{h}]", nesting_level=nesting_level)

    return width, height, DPI, DPI
