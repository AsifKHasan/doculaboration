#!/usr/bin/env python

import time
import yaml
import argparse
from pathlib import Path

from doc.doc_helper import DocHelper
from doc.doc_util import *
from helper.logger import *
from helper import logger


class DocFromJson(object):

    def __init__(self):
        self.start_time = int(round(time.time() * 1000))


    def run(self, config_file, json_file=None, nesting_level=0):
        # configuration
        # configuration
        config_service = ConfigService(config_file=config_file, nesting_level=0)

        # initialize GoogleServices
        google_services = GoogleServices(json_path=config_service._google_cred_json_path, nesting_level=0)


        # if gsheet name was provided as parameter, override the gsheet_list
        if gsheet:
            config_service._gsheet_list = [gsheet]


        self._CONFIG = yaml.load(open(self._config_path, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
        config_dir = self._config_path.parent

        logger.LOG_LEVEL = self._CONFIG.get("log-level", 0)

        # page specs
        page_spec_file = config_dir / 'page-specs.yml'
        self._CONFIG['page-specs'] = yaml.load(open(page_spec_file, 'r', encoding='utf-8'), Loader=yaml.FullLoader)

		# style specs
        style_spec_file = config_dir / 'style-specs.yml'
        if Path.exists(page_spec_file):
            self._CONFIG['style-specs'] = yaml.load(open(style_spec_file, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
        else:
            warn(f"No style-spec [{style_spec_file}]' found .. will not override any style")

        # if json name was provided as parameter, override the configuration
        if self._json:
            self._CONFIG["jsons"] = [self._json]

        self._CONFIG['dirs']['output-dir'] = config_dir / self._CONFIG['dirs']['output-dir']
        self._CONFIG['dirs']['temp-dir'] = self._CONFIG['dirs']['output-dir'] / 'tmp'
        if not Path.exists(self._CONFIG['dirs']['temp-dir']):
            Path.mkdir(self._CONFIG['dirs']['temp-dir'])

        self._CONFIG['dirs']['temp-dir'] = str(self._CONFIG['dirs']['temp-dir']).replace('\\', '/')

        self._CONFIG['files']['docx-template'] = config_dir / self._CONFIG['files']['docx-template']

        if not 'files' in self._CONFIG:
            self._CONFIG['files'] = {}


        # process jsons one by one
        for json in self._CONFIG['jsons']:
            self._CONFIG['files']['input-json'] = f"{self._CONFIG['dirs']['output-dir']}/{json}.json"

            # load json file
            with open(self._CONFIG['files']['input-json'], "r") as f:
                self._data = json.load(f)

            # doc-helper
            self._CONFIG['files']['output-docx'] = f"{self._CONFIG['dirs']['output-dir']}/{json}.docx"
            doc_helper = DocHelper(self._CONFIG)
            doc_helper.generate_and_save(self._data['sections'])

            if self._CONFIG['docx-related']['generate-pdf']:
                info(msg=f"generating pdf ..")
                pdf_start_time = int(round(time.time() * 1000))
                generate_pdf(self._CONFIG['files']['output-docx'], self._CONFIG['dirs']['output-dir'])
                self.end_time = int(round(time.time() * 1000))
                info(msg=f"generating pdf .. done {(self.end_time - pdf_start_time)/1000} seconds")

            # tear down
            self.end_time = int(round(time.time() * 1000))
            info(msg=f"script took {(self.end_time - self.start_time)/1000} seconds")

    def set_up(self):




if __name__ == '__main__':
	# construct the argument parse and parse the arguments
	ap = argparse.ArgumentParser()
	ap.add_argument("-c", "--config", required=True, help="configuration yml path")
	ap.add_argument("-j", "--json", required=False, help="json name to override json list provided in configuration")
	args = vars(ap.parse_args())

	generator = DocFromJson()
	generator.run(config_file=args["config"], json_file=args["json"])
