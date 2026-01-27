#!/usr/bin/env python
'''
'''

import yaml
from pathlib import Path

from helper.logger import *
from helper import logger

class ConfigService:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(ConfigService, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self, config_file=None, nesting_level=0):
        if self._initialized:
            return
        
        self._config_path = Path(config_file).resolve()
        self._config_dir = self._config_path.parent
        _config_dict = yaml.safe_load(open(self._config_path, 'r', encoding='utf-8'))

        self._log_level = _config_dict.get("log-level", 0)
        logger.LOG_LEVEL = self._log_level

        self._gsheet_list = _config_dict.get('gsheets', [])
        self._output_dir = Path(_config_dict.get('output-dir', '../../out')).resolve()
        self._google_cred_json_path = _config_dict.get('google-cred', None)
        self._autocrop_pdf_pages = _config_dict.get('autocrop-pdf-pages', False)

        self._temp_dir = self._output_dir / 'tmp'
        self._temp_dir.mkdir(parents=True, exist_ok=True)

        self._index_worksheet = _config_dict.get('index-worksheet', '-toc')
        self._gsheet_read_wait_seconds = _config_dict.get('gsheet-read-wait-seconds', 60)
        self._gsheet_read_try_count = _config_dict.get('gsheet-read-try-count', 3)

        self._initialized = True

