#!/usr/bin/env python3
'''
'''
import sys
import json
import time
import yaml
import datetime
import argparse
from pathlib import Path

from odt.odt_helper import OdtHelper
from odt.odt_util import *
from helper.logger import *
from helper import logger

class OdtFromJson(object):

	def __init__(self, config_path, json=None):
		self.start_time = int(round(time.time() * 1000))
		self._config_path = Path(config_path).resolve()
		self._data = {}
		self._json = json

	def run(self):
		self.set_up()
		# process jsons one by one
		for json in self._CONFIG['jsons']:
			self._CONFIG['files']['input-json'] = f"{self._CONFIG['dirs']['output-dir']}/{json}.json"
			self.load_json()

			# odt-helper
			self._CONFIG['files']['output-odt'] = f"{self._CONFIG['dirs']['output-dir']}/{json}.odt"
			odt_helper = OdtHelper(self._CONFIG)
			odt_helper.generate_and_save(self._data['sections'])

			if self._CONFIG['odt-related']['generate-pdf']:
				info(msg=f"generating pdf ..")
				pdf_start_time = int(round(time.time() * 1000))
				generate_pdf(self._CONFIG['files']['output-odt'], self._CONFIG['dirs']['output-dir'])
				self.end_time = int(round(time.time() * 1000))
				info(msg=f"generating pdf .. done {(self.end_time - pdf_start_time)/1000} seconds")

			self.tear_down()

	def set_up(self):
		# configuration
		self._CONFIG = yaml.load(open(self._config_path, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
		config_dir = self._config_path.parent

		logger.LOG_LEVEL = self._CONFIG.get('log-level', 0)

		# page specs
		page_spec_file = config_dir / 'page-specs.yml'
		self._CONFIG['page-specs'] = yaml.load(open(page_spec_file, 'r', encoding='utf-8'), Loader=yaml.FullLoader)

		# if json name was provided as parameter, override the configuration
		if self._json:
			self._CONFIG['jsons'] = [self._json]

		self._CONFIG['dirs']['output-dir'] = config_dir / self._CONFIG['dirs']['output-dir']
		self._CONFIG['dirs']['temp-dir'] = self._CONFIG['dirs']['output-dir'] / 'tmp'
		if not Path.exists(self._CONFIG['dirs']['temp-dir']):
			Path.mkdir(self._CONFIG['dirs']['temp-dir'])

		self._CONFIG['dirs']['temp-dir'] = str(self._CONFIG['dirs']['temp-dir']).replace('\\', '/')

		self._CONFIG['files']['odt-template'] = config_dir / self._CONFIG['files']['odt-template']

		if not 'files' in self._CONFIG:
			self._CONFIG['files'] = {}

	def load_json(self):
		with open(self._CONFIG['files']['input-json'], "r") as f:
			self._data = json.load(f)

	def tear_down(self):
		self.end_time = int(round(time.time() * 1000))
		debug(msg=f"script took {(self.end_time - self.start_time)/1000} seconds")

if __name__ == '__main__':
	# construct the argument parse and parse the arguments
	ap = argparse.ArgumentParser()
	ap.add_argument("-c", "--config", required=True, help="configuration yml path")
	ap.add_argument("-j", "--json", required=False, help="json name to override json list provided in configuration")
	args = vars(ap.parse_args())

	generator = OdtFromJson(args["config"], args["json"])
	generator.run()
