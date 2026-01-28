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

        self._google_cred_json_path = Path(_config_dict.get('google-cred', None)).resolve()
        self._json_list = _config_dict.get('jsons', [])
        self._data_dir = Path(_config_dict.get('data-dir', '../data')).resolve()
        self._output_dir = Path(_config_dict.get('output-dir', '../../out')).resolve()

        self._temp_dir = self._output_dir / 'tmp'
        self._temp_dir.mkdir(parents=True, exist_ok=True)

        self._docx_template = Path(_config_dict.get('docx-template', None)).resolve()
        self._generate_pdf = _config_dict.get('generate-pdf', True)

        # page specs
        page_spec_file = self._config_dir / 'page-specs.yml'
        self._page_specs = yaml.safe_load(open(page_spec_file, 'r', encoding='utf-8'))

        # font specs
        # font_spec_file = self._config_dir / 'font-specs.yml'
        # if Path.exists(font_spec_file):
        #     self._font_specs = yaml.safe_load(open(font_spec_file, 'r', encoding='utf-8'))
        # else:
        #     warn(f"No font-spec [{font_spec_file}]' found .. no fonts to register", nesting_level=nesting_level)

        # style specs
        style_spec_file = self._config_dir / 'style-specs.yml'
        if Path.exists(page_spec_file):
            self._style_specs = yaml.safe_load(open(style_spec_file, 'r', encoding='utf-8'))
        else:
            warn(f"No style-spec [{style_spec_file}]' found .. will not override any style", nesting_level=nesting_level)

        self._initialized = True

