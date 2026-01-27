#!/usr/bin/env python
'''
'''

from google.oauth2 import service_account
from googleapiclient.discovery import build

from helper.logger import *

class GoogleServices:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(GoogleServices, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self, json_path=None, nesting_level=0):
        if self._initialized:
            return
        
        if not json_path:
            raise ValueError("JSON key path must be provided for the first initialization.")

        # Define the scopes required
        self.scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]

        debug(f"authorizing with Google", nesting_level=0)
        # Load credentials from the service account JSON
        self._credential = service_account.Credentials.from_service_account_file(json_path, scopes=self.scopes)

        # Build and cache the services
        self._sheet_service = build('sheets', 'v4', credentials=self._credential)
        self._drive_service = build('drive', 'v3', credentials=self._credential)
        
        debug(f"authorized  with Google", nesting_level=0)
        self._initialized = True

    @property
    def sheets_api(self):
        return self._sheet_service

    @property
    def drive_api(self):
        return self._drive_service
