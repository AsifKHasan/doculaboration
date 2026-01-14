#!/usr/bin/env python
'''
'''

import io
import re
from pathlib import Path

import requests
import urllib3

from googleapiclient import errors
from googleapiclient.http import MediaIoBaseDownload

# import pygsheets
from PIL import Image

from helper.logger import *


# -------------------------------------------------------------------------------------------------------
# GSheet utilities  
# -------------------------------------------------------------------------------------------------------
''' get data from a gsheet
'''
def get_gsheet_data(google_service, gsheet, ranges=[], include_grid_data=True, nesting_level=0):
    # The spreadsheet to request.
    spreadsheet_id = gsheet.id

    request = google_service.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=ranges, includeGridData=include_grid_data)
    response = request.execute()

    return response


''' check whether a worksheet exists in a gsheet
'''
def worksheet_exists(sheet, ws_title, nesting_level=0):
    try:
        ws = sheet.worksheet('title', ws_title)
        return True
    except:
        warn(f"no worksheet ... [{ws_title}]", nesting_level=nesting_level)
        return False



# -------------------------------------------------------------------------------------------------------
# GDrive utilities  
# -------------------------------------------------------------------------------------------------------

''' get drive file metadata dict given a file id

'''
def drive_file_metadata(drive_service, file_id, nesting_level=0):
    file = drive_service.files().get(fileId=file_id,fields="id,name,mimeType").execute()
    return file


''' download media from drive given the id to a local path
'''
def download_media_from_dive(drive_service, file_id, local_path, nesting_level=0):
    request = drive_service.files().get_media(fileId=file_id,)

    fh = io.FileIO(local_path, "wb")
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()
        trace(f"Downloaded {int(status.progress() * 100)}%", nesting_level=nesting_level)

    return done


''' download a file from a drive url and return a dict
    {'file-name': file-name, 'file-type': file-type, 'file-path': local_path)}
    drive file id may have many forms like
    "https://drive.google.com/open?id=1rVmH-dHciYgPwJC0EFpHIXaZ_H1j5LDu"
    "https://drive.google.com/file/d/10bjxB_yjXgtaGZWJLVCQ28a7G_PPBLk-"
'''
def download_file_from_drive(drive_service, url, title, tmp_dir, nesting_level=0):
    file_url = url.strip()

    id = file_url.replace('https://drive.google.com/', '')
    # print(id)

    # see if it has something like 'id=' in it, then it will start after the pattern
    id = re.sub(r".*\?id=", "", id)

    # see if it has something like '/d/' in it, then it will start after the pattern
    id = re.sub(r".*/d/", "", id)

    # then it will be till end or till before the first '/'
    id = id.split('/')[0]

    # get file metadata
    file = drive_file_metadata(drive_service=drive_service, file_id=id, nesting_level=nesting_level+1)
    if not file:
        error(f"drive file [{id}] could not be accessed", nesting_level=nesting_level)
        return None
    
    if title:
        file_name = title
    else:
        file_name = file['name']
    
    file_type = file['mimeType']
    if not file_type in ALLOWED_MIME_TYPES:
        warn(f"drive url {url} is not a [{'/'.join(SUPPORTED_FILE_FORMATS)}], it is [{file_type}]", nesting_level=nesting_level)
        return None
    
    # determine file extension
    expected_extension = MIME_TYPE_TO_FILE_EXT_MAP.get(file_type, None)
    if expected_extension and not file_name.endswith(expected_extension):
        file_name = file_name + expected_extension

    local_path = f"{tmp_dir}/{file_name}"
    local_path = Path(local_path).resolve()

    # if the file is already in the local_path, we do not download it
    if Path(local_path).exists():
        trace(f"drive file existing   at: [{local_path}]", nesting_level=nesting_level)
        return {'file-name': file_name, 'file-type': file_type, 'file-path': str(local_path)}

    # finally download the file
    trace(f"downloading drive file id = [{id}]", nesting_level=nesting_level)
    try:
        download_media_from_dive(drive_service=drive_service, file_id=id, local_path=local_path, nesting_level=nesting_level+1)
        trace(f"drive file downloaded at: [{local_path}]", nesting_level=nesting_level)
        return {'file-name': file_name, 'file-type': file_type, 'file-path': str(local_path)}

    except:
        error(f"could not download : [{file_url}]", nesting_level=nesting_level)
        return None


''' copy drive file of a given id with a new name/title
    Args:
        service: Drive API service instance.
        origin_file_id: ID of the origin file to copy.
        copy_title: Title of the copy.

    Returns:
        The copied file if successful, None otherwise.
'''
def copy_drive_file(drive_service, origin_file_id, copy_title, nesting_level):
    copied_file = {'title': copy_title}
    try:
        return drive_service.files().copy(fileId=origin_file_id, body=copied_file).execute()
    except(errors.HttpError, error):
        error(f"An error occurred: {error}", nesting_level=nesting_level)
        return None


''' read content of a drive file from drive url
'''
def read_drive_file(drive_service, drive_url, nesting_level):
    url = drive_url.strip()

    id = url.replace('https://drive.google.com/file/d/', '')
    id = id.split('/')[0]
    # debug(f"drive file id to be read from is {id}", nesting_level=nesting_level)
    f = context['drive'].CreateFile({'id': id})
    if f['mimeType'] != 'text/plain':
        warn(f"drive url {url} mime-type is {f['mimeType']} which may not be readable as text", nesting_level=nesting_level)

    text = f.GetContentString()
    return text


''' get a drive file from title
'''
def get_drive_file(drive_service, drive_file_name, verbose=False, nesting_level=0):
    try:
        q = f"name = '{drive_file_name}'"
        files = []
        page_token = None
        while True:
            response = drive_service.files().list(q=q,
                                            spaces='drive',
                                            fields='nextPageToken, files(id, name, webViewLink, owners)',
                                            pageToken=page_token).execute()

            files.extend(response.get('files', []))
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

        # print(files)
        if len(files) > 0:
            if verbose:
                if len(files) > 1:
                    warn(f"multiple drive files found for [{drive_file_name}] ... ", nesting_level=nesting_level)

                for i, drive_file in enumerate(files, start=1):
                    full_path = get_file_hierarchy(drive_service=drive_service, file_id=drive_file['id'])
                    path_name = " -> ".join([item['name'] for item in reversed(full_path)])
                    owner = drive_file['owners'][0]['emailAddress']
                    if i == 1:
                        trace(f"{i} : owned by [{owner}] located in {path_name}", nesting_level=nesting_level)
                    else:
                        warn(f"{i} : owned by [{owner}] located in {path_name}", nesting_level=nesting_level)


            return files[0]

    except Exception as error:
        error(f"An error occurred: {error}", nesting_level=nesting_level)
        return None


''' Recursively retrieves the full parent folder hierarchy for a given file ID.

Args:
    file_id (str): The ID of the file to trace.
    service: The authenticated Google Drive API service object.

Returns:
    list: A list of dicts representing the path from the file up to the root.
            Example: [{'id': 'fileId', 'name': 'MyFile'}, {'id': 'parent1Id', 'name': 'Folder A'}, ...]
'''
def get_file_hierarchy(drive_service, file_id, nesting_level=0):
    hierarchy = []
    current_id = file_id

    # Use a loop to trace back the parents
    while current_id:
        try:
            # 1. Fetch file/folder metadata
            # We specifically request the 'id', 'name', and 'parents' fields
            file_metadata = drive_service.files().get(
                fileId=current_id, 
                fields='id, name, parents'
            ).execute()
            
            file_name = file_metadata.get('name')
            parent_ids = file_metadata.get('parents')
            
            # 2. Add the current item to the hierarchy list
            hierarchy.append({'id': current_id, 'name': file_name})

            # 3. Determine the next parent ID for the next iteration
            if parent_ids:
                # Drive can have multiple parents; we typically follow the first one
                # for a single path, or you could adapt this to trace all paths.
                current_id = parent_ids[0] 
            else:
                # No more parents means we have reached the root (My Drive)
                current_id = None

        except Exception as e:
            error(f"An error occurred while fetching ID {current_id}: {e}")
            break
    
    # The hierarchy is built from the file upward; return it in that order.
    return hierarchy



# -------------------------------------------------------------------------------------------------------
# Web utilities  
# -------------------------------------------------------------------------------------------------------
''' download a file from a web url and return a dict
    {'file-name': file-name, 'file-type': file-type, 'file-path': local_path)}
'''
def download_file_from_web(url, tmp_dir, nesting_level=0):
    file_url = url.strip()
    file_parts = file_url.split('.')
    if len(file_parts) > 1:
        file_ext = '.' + file_parts[-1]

    if not file_ext in SUPPORTED_FILE_FORMATS:
        error(f"url {file_url} is NOT a [{'/'.join(SUPPORTED_FILE_FORMATS)}] file", nesting_level=nesting_level)
        return None

    file_name = file_url.split('/')[-1].strip()
    file_type = FILE_EXT_TO_MIME_TYPE_MAP.get(file_ext, None)

    # download pdf in url into localpath
    try:
        local_path = f"{tmp_dir}/{file_name}"
        local_path = Path(local_path).resolve()
        # if the pdf is already in the local_path, we do not download it
        if Path(local_path).exists():
            trace(f"file existing   [{file_url}]", nesting_level=nesting_level)
            # pass
        else:
            file_data = requests.get(file_url).content
            with open(local_path, 'wb') as handler:
                handler.write(file_data)

            trace(f"file downloaded [{file_url}]", nesting_level=nesting_level)

        return {'file-name': file_name, 'file-type': file_type, 'file-path': str(local_path)}
    except:
        error(f"could not download : [{file_url}]", nesting_level=nesting_level)
        return None


''' read content from a web url
'''
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



# -------------------------------------------------------------------------------------------------------
# Image utilities  
# -------------------------------------------------------------------------------------------------------
''' get image metadata using Pillow
'''
def image_meta_pillow(im_path, nesting_level=0):
    im = Image.open(im_path)
    width, height = im.size

    if 'dpi' in im.info:
        dpi_x, dpi_y = im.info['dpi']
    else:
        dpi_x, dpi_y = DPI, DPI

    return width, height, dpi_x, dpi_y


''' from an image path return image params as a dict
    {'height': height, 'width': width, 'dpi': dpi, 'size': (size)}
'''
def image_params_from_image(local_path, row_height, mode, formula, formula_parts, nesting_level=0):
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
            im_dpi = (DPI, DPI)

        aspect_ratio = (im_width / im_height)
    except:
        warn(f"could not get dimesnion for image: [{local_path}]", nesting_level=nesting_level)
        return None

    # image link is based on row height
    if mode == 1:
        height = row_height
        width = int(height * aspect_ratio)
        # trace(f"adjusting image {local_path} at {width}x{height}-{im_dpi} based on row height {row_height}", nesting_level=nesting_level)
        return {'height': height, 'width': width, 'dpi': im_dpi, 'size': im.size}

    # image link is without height, width - use actual image size
    if mode == 3:
        # trace(f"keeping image {local_path} at {im_width}x{im_height}-{im_dpi} as-is", nesting_level=nesting_level)
        return {'height': im_height, 'width': im_width, 'dpi': im_dpi, 'size': im.size}

    # image link specifies height and width, use those
    if mode == 4 and len(formula_parts) == 4:
        # trace(f"image {local_path} at {im_width}x{im_height}-{im_dpi} size specified", nesting_level=nesting_level)
        return {'height': int(formula_parts[2]), 'width': int(formula_parts[3]), 'dpi': im_dpi, 'size': im.size}
    else:
        warn(f"image link does not specify height and width: [{formula}]", nesting_level=nesting_level)
        return None


''' from inside an IMAGE formula get the image a return a dict
    {'url': url, 'path': local_path, 'mode': mode}
    image_formula liiks like
    "http://documents.biasl.net/data/projects/rhd/filling-station-367x221.png", 3'\
    or this
    "http://documents.biasl.net/data/res/logo/rhd-logo-200x200.png", 4, 150, 150
'''
def download_image_from_formula(image_formula, tmp_dir, row_height, nesting_level=0):
    s = image_formula.replace('"', '').split(',')
    url, local_path, mode = None, None, None

    # the first item is url
    if len(s) >= 1:
        url = s[0]

        # localpath is the last term if it ends with png/jpg/gif/webp, if not
        url_splitted = url.split('/')
        if url_splitted[-1].endswith((tuple(SUPPORTED_FILE_FORMATS))):
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
                trace(f"image existing   at: [{local_path}]", nesting_level=nesting_level)
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

    # the second item is mode - can be 1, 3 or 4
    if len(s) >= 2:
        mode = int(s[1])
    else:
        warn(f"image link mode not found in formula: [{image_formula}]", nesting_level=nesting_level)
        return None

    # get the image dimensions
    image_params = image_params_from_image(local_path=local_path, row_height=row_height, mode=mode, formula=image_formula, formula_parts=s, nesting_level=nesting_level+1)
    if image_params:
        return {**{'url': url, 'path': str(local_path), 'mode': mode}, **image_params}
    else:
        return None


''' download an image from a web or drive url and return a dict
    {'file-path': file-path, 'file-type': file-type, 'image-height': height, 'image-width': width}
'''
def download_image(drive_service, url, title, tmp_dir, nesting_level=0):
    data = None
    if url.startswith('https://drive.google.com/'):
        data = download_file_from_drive(drive_service=drive_service, url=url, title=title, tmp_dir=tmp_dir, nesting_level=nesting_level)

    elif url.startswith('http'):
        # the file url is a normal web url
        data = download_file_from_web(url=url, tmp_dir=tmp_dir, nesting_level=nesting_level)

    else:
        warn(f"the url [{url}] is not a drive or web url", nesting_level=nesting_level)
        return None

    # if image, calculate dimensions
    if data['file-type'] in IMAGE_MIME_TYPES:
        file_path = data['file-path']
        width, height, dpi_x, dpi_y = image_meta_pillow(file_path, nesting_level=nesting_level)
        data['image-width'] = float(width / dpi_x)
        data['image-height'] = float(height / dpi_y)

    return data



# -------------------------------------------------------------------------------------------------------
# General utilities  
# -------------------------------------------------------------------------------------------------------

''' extract gsheet id from a drive url
    https://docs.google.com/spreadsheets/d/187T6rDV_I6ZwH1OlFYNdWN2OK05Z7ZxOQgqG0BhRbrE/edit?usp=drivesdk
'''
def gsheet_id_from_url(url, nesting_level=0):
    s = url[39:]
    splits = s.split('/')
    return splits[0]



''' from a HYPERLINK formula get the drive url and link name
'''
def get_gsheet_link(value, nesting_level=0):
    link_name, link_target = value, None

    m = re.match(r'=HYPERLINK\("(?P<link_url>.+)",\s*"(?P<link_title>.+)"\)', value, re.IGNORECASE)
    if m and m.group('link_url') is not None and m.group('link_title') is not None:
        # trace(f".. found a link to [{url}] at [{m.group('link_url')}]", nesting_level=nesting_level)
        link_name, link_target = m.group('link_title'), m.group('link_url')

    else:
        link_target = value


    return link_name, link_target



''' from a HYPERLINK formula get the worksheet id
'''
def get_worksheet_link(value, nesting_level=0):
    link_name = value

    # content can be a HYPERLINK/hyperlink to another worksheet
    m = re.match(r'=HYPERLINK\("#gid=(?P<ws_gid>.+)",\s*"(?P<ws_title>.+)"\)', value, re.IGNORECASE)
    if m and m.group('ws_gid') is not None and m.group('ws_title') is not None:
        # trace(m.group('ws_gid'), m.group('ws_title'), nesting_level=nesting_level)
        link_name = m.group('ws_title')

    return link_name



''' from a dict lookup a key, return one value if a cetain value is in the key, else return another value
    if look_up_value is passed None, do not check for value, just return he value
'''
def translate_dict_to_value(data_list, dict_obj, first_key, look_up_key='column', look_up_value=None, on_success=True, on_failure=False, nesting_level=0):
    if first_key not in dict_obj:
        # trace(f"[{first_key:<20}]: not found in dict .. defaulting to [{on_failure}]", nesting_level=nesting_level)
        return on_failure

    obj = dict_obj[first_key]

    if look_up_key not in obj:
        if 'value-if-missing' in obj:
            value_to_return = obj['value-if-missing']
        else:
            value_to_return = on_failure

        # trace(f"[{first_key:<20}]:[{look_up_key}] NOT found in dict .. defaulting to [{value_to_return}]", nesting_level=nesting_level)
        return value_to_return

    # get the value in the data_list
    value = data_list[obj[look_up_key]].strip()
    if look_up_value is None:
        # trace(f"[{first_key:<20}]: found .. returning '{value}'", nesting_level=nesting_level)
        return value

    if value == look_up_value:
        # trace(f"[{first_key:<20}]: value '{value}' matched with expected '{look_up_value}' ... returning [{on_success}]", nesting_level=nesting_level)
        return on_success

    else:
        # trace(f"[{first_key:<20}]: value '{value}' NOT matched with expected '{look_up_value}' ... returning [{on_failure}]", nesting_level=nesting_level)
        return on_failure



# -------------------------------------------------------------------------------------------------------
# data and globals 
# -------------------------------------------------------------------------------------------------------

SUPPORTED_FILE_FORMATS = ['.pdf', '.png', '.jpg', '.gif', '.webp']
IMAGE_FORMATS = ['.png', '.jpg', '.gif', '.webp']

ALLOWED_MIME_TYPES = ['application/pdf', 'image/png', 'image/jpeg', 'image/gif', 'image/webp']
IMAGE_MIME_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp']

FILE_EXT_TO_MIME_TYPE_MAP = {
    '.pdf': 'application/pdf', 
    '.png': 'image/png', 
    '.jpg': 'image/jpeg', 
    '.gif': 'image/gif', 
    '.webp': 'image/webp' 
}

MIME_TYPE_TO_FILE_EXT_MAP = {
    'application/pdf': '.pdf', 
    'image/png': '.png', 
    'image/jpeg': '.jpg', 
    'image/gif': '.gif', 
    'image/webp': '.webp'
}

DPI = 72

JPEG_QUALITY_DEFAULT = 90