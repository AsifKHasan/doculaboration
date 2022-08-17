#!/usr/bin/env python3

import sys
import pygsheets

import httplib2

from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

from helper.logger import *
from helper.gsheet.gsheet_util import *
from helper.gsheet.gsheet_reader import *

class GsheetHelper(object):

    __instance = None

    ''' class constructor
    '''
    def __new__(cls):
        # we only need one singeton instance of this class
        if GsheetHelper.__instance is None:
            GsheetHelper.__instance = object.__new__(cls)

        return GsheetHelper.__instance



    ''' initialize the helper
    '''
    def init(self, config):
        # as we go further we put everything inside a single dict _context
        self._context = {}

        info(f"authorizing with Google")

        _G = pygsheets.authorize(service_account_file=config['files']['google-cred'])
        self._context['_G'] = _G

        credentials = ServiceAccountCredentials.from_json_keyfile_name(config['files']['google-cred'], scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'])
        credentials.authorize(httplib2.Http())

        self._context['service'] = discovery.build('sheets', 'v4', credentials=credentials)

        gauth = GoogleAuth()
        gauth.credentials = credentials

        self._context['drive'] = GoogleDrive(gauth)
        self._context['tmp-dir'] = config['dirs']['temp-dir']
        self._context['index-worksheet'] = config['index-worksheet']
        self._context['gsheet-read-wait-seconds'] = config['gsheet-read-wait-seconds']
        self._context['gsheet-read-try-count'] = config['gsheet-read-try-count']

        info(f"authorized  with Google")


    ''' read the gsheet
    '''
    def read_gsheet(self, gsheet_title, gsheet_url=None, parent=None):
        wait_for = self._context['gsheet-read-wait-seconds']
        try_count = self._context['gsheet-read-try-count']
        gsheet = None
        for i in range(0, try_count):
            try:
                if gsheet_url:
                    info(f"opening gsheet with url : {gsheet_url}")
                    gsheet = self._context['_G'].open_by_url(gsheet_url)
                    info(f"opened  gsheet with url : {gsheet_url}")
                else:
                    info(f"opening gsheet : {gsheet_title}")
                    gsheet = self._context['_G'].open(gsheet_title)
                    info(f"opened  gsheet : {gsheet_title}")

                # optimization - read the full gsheet
                info(f"reading gsheet : {gsheet_title}")
                self._context['gsheet-data'] = self.get_gsheet_data(gsheet)
                info(f"read    gsheet : {gsheet_title}")
                # return self._context['gsheet-data']
                
                break

            except:
                warn(f"gsheet {gsheet_title} read request (attempt {i}) failed, waiting for {wait_for} seconds before trying again")
                time.sleep(float(wait_for))

        if gsheet is None:
            error('gsheet read request failed, quiting')
            sys.exit(1)

        return process_gsheet(self._context, gsheet, parent)



    ''' get data from the gsheet
    '''
    def get_gsheet_data(self, gsheet):
        # The spreadsheet to request.
        spreadsheet_id = gsheet.id

        # The ranges to retrieve from the spreadsheet
        ranges = []

        # This parameter is ignored if a field mask was set in the request
        include_grid_data = True

        request = self._context['service'].spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=ranges, includeGridData=include_grid_data)
        response = request.execute()

        # make a dictionary key'ed by worksheet_name
        response = {sheet['properties']['title']: sheet for sheet in response['sheets']}

        return response

