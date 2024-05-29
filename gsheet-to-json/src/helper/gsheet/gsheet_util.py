#!/usr/bin/env python3

import re
from pathlib import Path

import requests
import urllib3

import pygsheets
from PIL import Image

from helper.logger import *


''' get data from the gsheet
'''
def get_gsheet_data(google_service, gsheet, ranges=[], include_grid_data=True):
    # The spreadsheet to request.
    spreadsheet_id = gsheet.id

    request = google_service.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=ranges, includeGridData=include_grid_data)
    response = request.execute()

    return response



'''
    https://docs.google.com/spreadsheets/d/187T6rDV_I6ZwH1OlFYNdWN2OK05Z7ZxOQgqG0BhRbrE/edit?usp=drivesdk
'''
def gsheet_id_from_url(url, nesting_level):
    s = url[39:]
    splits = s.split('/')
    return splits[0]



def get_gsheet_link(value, nesting_level=0):
    link_name, link_target = value, None

    m = re.match('=HYPERLINK\("(?P<link_url>.+)",\s*"(?P<link_title>.+)"\)', value, re.IGNORECASE)
    if m and m.group('link_url') is not None and m.group('link_title') is not None:
        # debug(f".. found a link to [{url}] at [{m.group('link_url')}]", nesting_level=nesting_level)
        link_name, link_target = m.group('link_title'), m.group('link_url')

    else:
        link_target = value


    return link_name, link_target



def get_worksheet_link(value, nesting_level=0):
    link_name = value

    # content can be a HYPERLINK/hyperlink to another worksheet
    m = re.match('=HYPERLINK\("#gid=(?P<ws_gid>.+)",\s*"(?P<ws_title>.+)"\)', value, re.IGNORECASE)
    if m and m.group('ws_gid') is not None and m.group('ws_title') is not None:
        # debug(m.group('ws_gid'), m.group('ws_title'), nesting_level=nesting_level)
        link_name = m.group('ws_title')

    return link_name



def worksheet_exists(sheet, ws_title, nesting_level=0):
    try:
        ws = sheet.worksheet('title', ws_title)
        return True
    except:
        warn(f"no worksheet ... [{ws_title}]", nesting_level=nesting_level)
        return False



def download_image_from_formula(image_formula, tmp_dir, row_height, nesting_level=0):
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

        # localpath is the last term if it ends with png/jpg/gif/webp, if not
        url_splitted = url.split('/')
        if url_splitted[-1].endswith('.png') or url_splitted[-1].endswith('.jpg') or url_splitted[-1].endswith('.gif') or url_splitted[-1].endswith('.webp'):
            local_path = f"{tmp_dir}/{url_splitted[-1]}"

        # if it is owncloud, (https://storage.brilliant.com.bd/s/IPO46mdbcetahMf/download) we use the penaltimate term and append a .png
        elif len(url_splitted) >= 6 and 'storage.brilliant.com.bd' in url_splitted[2]:
            local_path = f"{tmp_dir}/{url_splitted[-2]}.png"

        else:
            warn(f"url pattern unknown for file: {url}", nesting_level=nesting_level)
            return None

        local_path = Path(local_path).resolve()

        # download image in url into localpath
        try:
            # if the image is already in the local_path, we do not download it
            if Path(local_path).exists():
                debug(f"image existing   at: [{local_path}]", nesting_level=nesting_level)
            else:
                response = requests.get(url, stream=True)
                if not response.ok:
                    warn(f"{response} could not download image: [{url}]", nesting_level=nesting_level)
                    return None

                with open(local_path, 'wb') as handle:
                    for block in response.iter_content(512):
                        if not block:
                            break

                        handle.write(block)

                debug(f"image downloaded at: [{local_path}]", nesting_level=nesting_level)

        except Exception as err:
            warn(f"could not download : [{url}]", nesting_level=nesting_level)
            print(Exception, err)
            return None

    # get the image dimensions
    try:
        im = Image.open(local_path)
        im_width, im_height = im.size
        if 'dpi' in im.info:
            # dpi values are of type IFDRational which is not JSON serializable, cast them to float
            im_dpi_x, im_dpi_y = im.info['dpi']
            im_dpi_x = float(im_dpi_x)
            im_dpi_y = float(im_dpi_y)
            im_dpi = (im_dpi_x, im_dpi_y)
        else:
            im_dpi = (96, 96)

        aspect_ratio = (im_width / im_height)
    except:
        warn(f"could not get dimesnion for image: [{local_path}]", nesting_level=nesting_level)
        return None

    # the second item is mode - can be 1, 3 or 4
    if len(s) >= 2:
        mode = int(s[1])
    else:
        warn(f"image link mode not found in formula: [{image_formula}]", nesting_level=nesting_level)
        return None

    # image link is based on row height
    if mode == 1:
        height = row_height
        width = int(height * aspect_ratio)
        # info(f"adjusting image {local_path} at {width}x{height}-{im_dpi} based on row height {row_height}", nesting_level=nesting_level)
        return {'url': url, 'path': str(local_path), 'height': height, 'width': width, 'dpi': im_dpi, 'size': im.size, 'mode': mode}

    # image link is without height, width - use actual image size
    if mode == 3:
        # info(f"keeping image {local_path} at {im_width}x{im_height}-{im_dpi} as-is", nesting_level=nesting_level)
        return {'url': url, 'path': str(local_path), 'height': im_height, 'width': im_width, 'dpi': im_dpi, 'size': im.size, 'mode': mode}

    # image link specifies height and width, use those
    if mode == 4 and len(s) == 4:
        # info(f"image {local_path} at {im_width}x{im_height}-{im_dpi} size specified", nesting_level=nesting_level)
        return {'url': url, 'path': str(local_path), 'height': int(s[2]), 'width': int(s[3]), 'dpi': im_dpi, 'size': im.size, 'mode': mode}
    else:
        warn(f"image link does not specify height and width: [{image_formula}]", nesting_level=nesting_level)
        return None



def download_file_from_web(url, tmp_dir, nesting_level=0):
    file_url = url.strip()
    file_ext = file_url[-4:]
    if not file_ext in ['.pdf', '.png', '.jpg', '.gif', '.webp']:
        error(f"url {file_url} is NOT a pdf/png/jpg/gif/webp file", nesting_level=nesting_level)
        return None

    file_name = file_url.split('/')[-1].strip()
    if file_ext == '.pdf':
        file_type = 'application/pdf'

    elif file_ext == '.png':
        file_type = 'image/png'
    
    elif file_ext == '.jpg':
        file_type = 'image/jpeg'

    elif file_ext == '.gif':
        file_type = 'image/gif'

    elif file_ext == '.webp':
        file_type = 'image/webp'

    # download pdf in url into localpath
    try:
        local_path = f"{tmp_dir}/{file_name}"
        local_path = Path(local_path).resolve()
        # if the pdf is already in the local_path, we do not download it
        if Path(local_path).exists():
            debug(f"file existing   [{file_url}]", nesting_level=nesting_level)
            # pass
        else:
            file_data = requests.get(file_url).content
            with open(local_path, 'wb') as handler:
                handler.write(file_data)

            debug(f"file downloaded [{file_url}]", nesting_level=nesting_level)

        return {'file-name': file_name, 'file-type': file_type, 'file-path': str(local_path)}
    except:
        error(f"could not download : [{file_url}]", nesting_level=nesting_level)
        return None



def download_file_from_drive(url, tmp_dir, drive, nesting_level=0):
    file_url = url.strip()

    id = file_url.replace('https://drive.google.com/file/d/', '')
    id = id.split('/')[0]
    debug(f"downloading drive file id = [{id}]", nesting_level=nesting_level)
    f = drive.CreateFile({'id': id})
    file_name = f['title']
    file_type = f['mimeType']
    if not file_type in ['application/pdf', 'image/png', 'image/jpeg', 'image/gif', 'image/webp']:
        warn(f"drive url {url} is not a pdf/png/jpg/gif/webp, it is [{f['mimeType']}]", nesting_level=nesting_level)
        return None

    if file_type == 'application/pdf' and not file_name.endswith('.pdf'):
        file_name = file_name + '.pdf'

    if file_type == 'image/png' and not file_name.endswith('.png'):
        file_name = file_name + '.png'

    if file_type == 'image/jpeg' and not file_name.endswith('.jpg'):
        file_name = file_name + '.jpg'

    if file_type == 'image/gif' and not file_name.endswith('.gif'):
        file_name = file_name + '.gif'

    if file_type == 'image/webp' and not file_name.endswith('.webp'):
        file_name = file_name + '.webp'

    try:
        local_path = f"{tmp_dir}/{file_name}"
        local_path = Path(local_path).resolve()
        # if the pdf is already in the local_path, we do not download it
        if Path(local_path).exists():
            debug(f"drive file existing   at: [{local_path}]", nesting_level=nesting_level)
            # pass            
        else:
            f.GetContentFile(local_path)
            debug(f"drive file downloaded at: [{local_path}]", nesting_level=nesting_level)

        return {'file-name': file_name, 'file-type': file_type, 'file-path': str(local_path)}
    except:
        error(f"could not download : [{file_url}]", nesting_level=nesting_level)
        return None



def read_web_content(web_url, nesting_level=0):
    url = web_url.strip()

    # read content from url
    try:
        http = urllib3.PoolManager()
        response = http.request('GET', url)
        text = response.data.decode('utf-8')
        return text
    except:
        error(f"could not read content from url: [{web_url}]", nesting_level=nesting_level)
        return None



def download_image(url, tmp_dir, nesting_level=0):
    data = None
    if url.startswith('http'):
        # the file url is a normal web url
        data = download_file_from_web(url=url, tmp_dir=tmp_dir)

    else:
        warn(f"the url [{url}] is not a web url", nesting_level=nesting_level)

    return data
