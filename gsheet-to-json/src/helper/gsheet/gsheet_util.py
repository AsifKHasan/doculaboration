#!/usr/bin/env python3

import re
import os.path
from os import path

import urllib3

import pygsheets
import urllib.request
from PIL import Image

from helper.logger import *

COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

def worksheet_exists(sheet, ws_title):
    try:
        ws = sheet.worksheet('title', ws_title)
        return True
    except:
        warn('No worksheet ... {0}'.format(ws_title))
        return False

def hex_to_rgba(color_hex):
    color = tuple(int(color_hex[i:i+2], 16) for i in (0, 2, 4, 6))
    return {'red': color[0]/255, 'green': color[1]/255, 'blue': color[2]/255, 'alpha': color[3]/255}

def range_object(ws, start_row, end_row, start_col, end_col):
    return {'sheetId': ws.id, 'startRowIndex': start_row, 'endRowIndex': end_row, 'startColumnIndex': start_col, 'endColumnIndex': end_col}

def dimension_range_object(ws, dimension, startIndex, endIndex):
    return {'sheetId': ws.id, 'dimension': dimension, 'startIndex': startIndex, 'endIndex': endIndex}

def delete_dimension_request(sheet, ws, dimension, startIndex, endIndex):
    return {
        'deleteDimension': {
            'range': dimension_range_object(ws, dimension, startIndex, endIndex)
        }
    }

def delete_dimension(sheet, ws, dimension, startIndex, endIndex):
    sheet.custom_request(delete_dimension_request(sheet, ws, dimension, startIndex, endIndex), None)

def resize_dimension_request(sheet, ws, dimension, start_index, size):
    fields = []
    fields.append('pixelSize')
    return {
        'updateDimensionProperties' : {
            'range' : dimension_range_object(ws, dimension, start_index, start_index + 1),
            'properties': {'pixelSize': size},
            'fields': ','.join(fields)
        }
    }

def resize_column(sheet, ws, column_index, width):
    sheet.custom_request(resize_dimension_request(sheet, ws, 'COLUMNS', column_index, width), None)

def freeze_request(sheet, ws, rows, cols):
    return {'updateSheetProperties': {'properties': {'sheetId': ws.id, 'gridProperties': {'frozenRowCount': rows, 'frozenColumnCount': cols}}, 'fields': 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'}}

def freeze(sheet, ws, rows, cols):
    sheet.custom_request(freeze_request(sheet, ws, rows, cols), None)

def format_range_request(sheet, ws, start_row, end_row, start_col, end_col, user_entered_value=None, font_family=None, font_size=None, bold=None, fg_color=None, bg_color=None, wrap_strategy=None, valign=None, halign=None, number_format=None, borders=None):
    fields = []
    if user_entered_value: fields.append('userEnteredValue.stringValue')

    if font_family: fields.append('userEnteredFormat.textFormat.fontFamily')
    if font_size: fields.append('userEnteredFormat.textFormat.fontSize')
    if bold: fields.append('userEnteredFormat.textFormat.bold')
    if fg_color: fields.append('userEnteredFormat.textFormat.foregroundColor')

    if bg_color: fields.append('userEnteredFormat.backgroundColor')
    if valign: fields.append('userEnteredFormat.verticalAlignment')
    if halign: fields.append('userEnteredFormat.horizontalAlignment')
    if wrap_strategy: fields.append('userEnteredFormat.wrapStrategy')
    if number_format: fields.append('userEnteredFormat.numberFormat')
    if borders: fields.append('userEnteredFormat.borders')

    return {
      'repeatCell': {
        'range': range_object(ws, start_row, end_row, start_col, end_col),
        'cell': {
          'userEnteredValue': {
            'stringValue': user_entered_value
            },
          'userEnteredFormat': {
            'verticalAlignment': valign,
            'horizontalAlignment': halign,
            'wrapStrategy': wrap_strategy,
            'numberFormat': number_format,
            'backgroundColor': None if bg_color is None else hex_to_rgba(bg_color),
            'borders': borders,
            'textFormat': {
              'foregroundColor': None if fg_color is None else hex_to_rgba(fg_color),
              'fontFamily': font_family,
              'fontSize': font_size,
              'bold': bold
            }
          }
        },
        'fields': ','.join(fields)
      }
    }

def format_range(sheet, ws, start_row, end_row, start_col, end_col, user_entered_value=None, font_family=None, font_size=None, bold=None, fg_color=None, bg_color=None, wrap_strategy=None, valign=None, halign=None, number_format=None, borders=None):
    sheet.custom_request(format_range_request(sheet, ws, start_row, end_row, start_col, end_col, user_entered_value, font_family, font_size, bold, fg_color, bg_color, wrap_strategy, valign, halign, number_format, borders), None)

def border_range_request(sheet, ws, start_row, end_row, start_col, end_col, borders=None):
    return {
      'updateBorders': {
        'range': range_object(ws, start_row, end_row, start_col, end_col),
        'top': borders['top'],
        'right': borders['right'],
        'bottom': borders['bottom'],
        'left': borders['left']
      }
    }

def border_range(sheet, ws, start_row, end_row, start_col, end_col, borders=None):
    sheet.custom_request(border_range_request(sheet, ws, start_row, end_row, start_col, end_col, borders), None)

def merge_cells_request(sheet, ws, start_row, end_row, start_col, end_col):
    return {
      'mergeCells': {
        'range': range_object(ws, start_row, end_row, start_col, end_col),
        'mergeType': 'MERGE_ALL'
      }
    }

def merge_cells(sheet, ws, start_row, end_row, start_col, end_col):
    sheet.custom_request(merge_cells_request(sheet, ws, start_row, end_row, start_col, end_col), None)

def init_worksheet(sheet, ws_title, num_rows=100, num_cols=4, frozen_rows=2, frozen_cols=0, col_def=None):
    # if worksheet exists, delete it and create again
    try:
        ws = sheet.worksheet_by_title(ws_title)
        sheet.del_worksheet(ws)
    except pygsheets.exceptions.WorksheetNotFound:
        pass

    ws = sheet.add_worksheet(ws_title, rows=num_rows, cols=num_cols)

    # freeze first row
    freeze(sheet, ws, frozen_rows, frozen_cols)

    # valign, font etc. for full worksheet
    format_range(sheet, ws, 0, num_rows, 0, num_cols, 'Calibri', 10, bold=False, fg_color='00000000', bg_color=None, wrap_strategy='WRAP', valign='MIDDLE', halign=None, number_format=None)

    # column size, alignment, format
    for col_letter, col_data in col_def.items():
        ws.adjust_column_width(start=col_data['idx'], end=None, pixel_size=col_data['size'])
        format_range(sheet, ws, 0, num_rows, col_data['idx'], col_data['idx'] + 1, None, None, bold=None, fg_color=None, bg_color=None, wrap_strategy=None, valign=None, halign=col_data['halign'], number_format=col_data['numberFormat'])

    return ws

def download_image(image_formula, tmp_dir, row_height):
    '''
        image_formula liiks like
        "http://documents.biasl.net/data/projects/rhd/filling-station-367x221.png", 3'\
        or this
        "http://documents.biasl.net/data/res/logo/rhd-logo-200x200.png", 4, 150, 150
    '''
    s = image_formula.replace('"', '').split(',')
    url, local_path, mode = None, None, None

    # the first item is url
    if len(s) >= 1:
        url = s[0]

        # localpath is the last term if it ends with png/jpg/gif, if not
        url_splitted = url.split('/')
        if url_splitted[-1].endswith('.png') or url_splitted[-1].endswith('.jpg') or url_splitted[-1].endswith('.gif'):
            local_path = '{0}/{1}'.format(tmp_dir, url_splitted[-1])

        # if it is owncloud, (https://storage.brilliant.com.bd/s/IPO46mdbcetahMf/download) we use the penaltimate term and append a .png
        elif len(url_splitted) >= 6 and 'storage.brilliant.com.bd' in url_splitted[2]:
                local_path = '{0}/{1}.png'.format(tmp_dir, url_splitted[-2])

        else:
            warn('.... url pattern unknown for file: {0}'.format(url))
            return None

        # download image in url into localpath
        try:
            # if the image is already in the local_path, we do not download it
            if path.exists(local_path):
                pass
            else:
                urllib.request.urlretrieve(url, local_path)
        except:
            warn('.... could not download file: {0}'.format(url))
            return None

    # get the image dimensions
    try:
        im = Image.open(local_path)
        im_width, im_height = im.size
        if 'dpi' in im.info:
            im_dpi = im.info['dpi']
        else:
            im_dpi = (96, 96)

        aspect_ratio = (im_width / im_height)
    except:
        warn('.... could not get dimesnion for image: {0}'.format(local_path))
        return None

    # the second item is mode - can be 1, 3 or 4
    if len(s) >= 2:
        mode = int(s[1])
    else:
        warn('.... image link mode not found in formula: {0}'.format(image_formula))
        return None

    # image link is based on row height
    if mode == 1:
        height = row_height
        width = int(height * aspect_ratio)
        info('.... adjusting image {0} at {1}x{2}-{3} based on row height {4}'.format(local_path, width, height, im_dpi, row_height))
        return {'url': url, 'path': local_path, 'height': height, 'width': width, 'dpi': im_dpi}

    # image link is without height, width - use actual image size
    if mode == 3:
        info('.... keeping image {0} at {1}x{2}-{3} as-is'.format(local_path, im_width, im_height, im_dpi))
        return {'url': url, 'path': local_path, 'height': im_height, 'width': im_width, 'dpi': im_dpi}

    # image link specifies height and width, use those
    if mode == 4 and len(s) == 4:
        info('.... image {0} at {1}x{2}-{3} size specified'.format(local_path, im_width, im_height, im_dpi))
        return {'url': url, 'path': local_path, 'height': int(s[2]), 'width': int(s[3]), 'dpi': im_dpi}
    else:
        info('.... image link does not specify height and width: {0}'.format(image_formula))
        return None

def download_pdf_from_web(url, tmp_dir):
    pdf_url = url.strip()
    if pdf_url[-4:] != '.pdf':
        error('.... url {0} is NOT a pdf file'.format(pdf_url))
        return None

    pdf_name = pdf_url.split('/')[-1].strip()

    # download pdf in url into localpath
    try:
        local_path = '{0}/{1}'.format(tmp_dir, pdf_name)
        # if the pdf is already in the local_path, we do not download it
        if path.exists(local_path):
            pass
        else:
            urllib.request.urlretrieve(pdf_url, local_path)
            info('.... {0} downloaded at: {1}'.format(pdf_url, local_path))

        return {'pdf_name': pdf_name, 'pdf_path': local_path}
    except:
        error('.... could not download pdf: {0}'.format(pdf_url))
        return None

def download_pdf_from_drive(url, tmp_dir, drive):
    pdf_url = url.strip()

    id = pdf_url.replace('https://drive.google.com/file/d/', '')
    id = id.split('/')[0]
    info('drive file id to be downloaded is {0}'.format(id))
    f = drive.CreateFile({'id': id})
    if f['mimeType'] != 'application/pdf':
        warn('drive url {0} is not a pdf'.format(url))
        return None

    pdf_name = f['title']
    if not pdf_name.endswith('.pdf'):
        pdf_name = pdf_name + '.pdf'

    try:
        local_path = '{0}/{1}'.format(tmp_dir, pdf_name)
        # if the pdf is already in the local_path, we do not download it
        if path.exists(local_path):
            pass
        else:
            f.GetContentFile(local_path)
            info('.... {0} downloaded at: {1}'.format(url, local_path))

        return {'pdf_name': pdf_name, 'pdf_path': local_path}
    except:
        error('.... could not download pdf: {0}'.format(pdf_url))
        return None

def read_web_content(web_url):
    url = web_url.strip()

    # read content from url
    try:
        http = urllib3.PoolManager()
        response = http.request('GET', url)
        text = response.data.decode('utf-8')
        return text
    except:
        error('.... could not read content from url: {0}'.format(web_url))
        return None
