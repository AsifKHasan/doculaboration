#!/usr/bin/env python3
'''
'''
import sys
import json
import importlib
import time
import yaml
import datetime
import argparse
import pprint
from pathlib import Path

from helper.logger import *
from helper.gsheet.gsheet_helper import GsheetHelper

class JsonFromGsheet(object):

	def __init__(self, config_path):
		self.start_time = int(round(time.time() * 1000))
		self._config_path = Path(config_path).resolve()
		self._data = {}

	def run(self):
		self.set_up()
		# process gsheets one by one
		for gsheet in self._CONFIG['gsheets']:
			self._data = self._gsheethelper.process_gsheet(gsheet)

			self._CONFIG['files']['output-json'] = '{0}/{1}.json'.format(self._CONFIG['dirs']['output-dir'], gsheet)
			self.save_json()

			self.tear_down()

	def set_up(self):
		# configuration
		self._CONFIG = yaml.load(open(self._config_path, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
		config_dir = self._config_path.parent

		self._CONFIG['dirs']['output-dir'] = config_dir / self._CONFIG['dirs']['output-dir']
		self._CONFIG['dirs']['temp-dir'] = self._CONFIG['dirs']['output-dir'] / 'tmp'
		if not Path.exists(self._CONFIG['dirs']['temp-dir']):
			Path.makedir(self._CONFIG['dirs']['temp-dir'])

		self._CONFIG['files']['google-cred'] = config_dir / self._CONFIG['files']['google-cred']

		# gsheet-helper
		self._gsheethelper = GsheetHelper()
		self._gsheethelper.init(self._CONFIG)

	def save_json(self):
		with open(self._CONFIG['files']['output-json'], "w") as f:
			f.write(json.dumps(self._data, sort_keys=False, indent=4))

	def tear_down(self):
		self.end_time = int(round(time.time() * 1000))
		debug("Script took {} seconds".format((self.end_time - self.start_time)/1000))

if __name__ == '__main__':
	# construct the argument parse and parse the arguments
	ap = argparse.ArgumentParser()
	ap.add_argument("-c", "--config", required=True, help="configuration yml path")
	args = vars(ap.parse_args())

	generator = JsonFromGsheet(args["config"])
	generator.run()
