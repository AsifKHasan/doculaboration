#!/usr/bin/env python3

from helper.logger import *

from googleapiclient import errors

def copy_drive_file(service, origin_file_id, copy_title):
    """
    Copy an existing file.

    Args:
        service: Drive API service instance.
        origin_file_id: ID of the origin file to copy.
        copy_title: Title of the copy.

    Returns:
        The copied file if successful, None otherwise.
    """
    copied_file = {'title': copy_title}
    try:
        return service.files().copy(fileId=origin_file_id, body=copied_file).execute()
    except(errors.HttpError, error):
        error(f"An error occurred: {error}")
        return None


def download_drive_file(param, destination, context):
    f = context['drive'].CreateFile(param)
    f.GetContentFile(destination)


def read_drive_file(drive_url, context):
    url = drive_url.strip()

    id = url.replace('https://drive.google.com/file/d/', '')
    id = id.split('/')[0]
    # debug(f"drive file id to be read from is {id}")
    f = context['drive'].CreateFile({'id': id})
    if f['mimeType'] != 'text/plain':
        warn(f"drive url {url} mime-type is {f['mimeType']} which may not be readable as text")

    text = f.GetContentString()
    return text
