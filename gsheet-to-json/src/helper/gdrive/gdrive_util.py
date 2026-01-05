#!/usr/bin/env python

from helper.logger import *

from googleapiclient import errors

def copy_drive_file(service, origin_file_id, copy_title, nesting_level):
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
        error(f"An error occurred: {error}", nesting_level=nesting_level)
        return None


def download_drive_file(param, destination, context, nesting_level):
    f = context['drive'].CreateFile(param)
    f.GetContentFile(destination)


def read_drive_file(drive_url, context, nesting_level):
    url = drive_url.strip()

    id = url.replace('https://drive.google.com/file/d/', '')
    id = id.split('/')[0]
    # debug(f"drive file id to be read from is {id}", nesting_level=nesting_level)
    f = context['drive'].CreateFile({'id': id})
    if f['mimeType'] != 'text/plain':
        warn(f"drive url {url} mime-type is {f['mimeType']} which may not be readable as text", nesting_level=nesting_level)

    text = f.GetContentString()
    return text


''' get a drive file
'''
def get_drive_file(service, drive_file_name, verbose=False, nesting_level=0):
    try:
        q = f"name = '{drive_file_name}'"
        files = []
        page_token = None
        while True:
            response = service.files().list(q=q,
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
                    full_path = get_file_hierarchy(service=service, file_id=drive_file['id'])
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


'''
Recursively retrieves the full parent folder hierarchy for a given file ID.

Args:
    file_id (str): The ID of the file to trace.
    service: The authenticated Google Drive API service object.

Returns:
    list: A list of dicts representing the path from the file up to the root.
            Example: [{'id': 'fileId', 'name': 'MyFile'}, {'id': 'parent1Id', 'name': 'Folder A'}, ...]
'''
def get_file_hierarchy(service, file_id, nesting_level=0):
    hierarchy = []
    current_id = file_id

    # Use a loop to trace back the parents
    while current_id:
        try:
            # 1. Fetch file/folder metadata
            # We specifically request the 'id', 'name', and 'parents' fields
            file_metadata = service.files().get(
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
