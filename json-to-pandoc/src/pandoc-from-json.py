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

from helper.logger import *
from helper.pandoc.pandoc_helper import Pandoc
from helper.pandoc.pandoc_util import *

class PandocFromJson(object):

	def __init__(self, config_path, json=None):
		self.start_time = int(round(time.time() * 1000))
		self._config_path = Path(config_path).resolve()
		self._data = {}
		self._json = json

	def run(self):
		self.set_up()
		# process jsons one by one
		for json in self._CONFIG['jsons']:
			self._CONFIG['files']['input-json'] = '{0}/{1}.json'.format(self._CONFIG['dirs']['output-dir'], json)
			self.load_json()

			# pandoc-helper
			self._CONFIG['files']['output-pandoc'] = '{0}/{1}.mkd'.format(self._CONFIG['dirs']['output-dir'], json)
			pandoc = Pandoc()
			pandoc.generate_pandoc(self._data['sections'], self._CONFIG, self._CONFIG['files']['pandoc-styles'], self._CONFIG['files']['document-header'], self._CONFIG['files']['output-pandoc'])
			self.tear_down()

	def set_up(self):
		# configuration
		self._CONFIG = yaml.load(open(self._config_path, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
		config_dir = self._config_path.parent

		# if json name was provided as parameter, override the configuration
		if self._json:
			self._CONFIG['jsons'] = [self._json]

		self._CONFIG['dirs']['output-dir'] = config_dir / self._CONFIG['dirs']['output-dir']
		self._CONFIG['dirs']['temp-dir'] = self._CONFIG['dirs']['output-dir'] / 'tmp'
		if not Path.exists(self._CONFIG['dirs']['temp-dir']):
			Path.mkdir(self._CONFIG['dirs']['temp-dir'])

		self._CONFIG['dirs']['temp-dir'] = str(self._CONFIG['dirs']['temp-dir']).replace('\\', '/')

		self._CONFIG['files']['pandoc-styles'] = config_dir / self._CONFIG['files']['pandoc-styles']
		self._CONFIG['files']['document-header'] = config_dir / self._CONFIG['files']['document-header']

		if not 'files' in self._CONFIG:
			self._CONFIG['files'] = {}

	def load_json(self):
		with open(self._CONFIG['files']['input-json'], "r") as f:
			self._data = json.load(f)

	def tear_down(self):
		self.end_time = int(round(time.time() * 1000))
		debug("Script took {} seconds".format((self.end_time - self.start_time)/1000))

if __name__ == '__main__':
	# construct the argument parse and parse the arguments
	ap = argparse.ArgumentParser()
	ap.add_argument("-c", "--config", required=True, help="configuration yml path")
	ap.add_argument("-j", "--json", required=False, help="json name to override json list provided in configuration")
	args = vars(ap.parse_args())

	generator = PandocFromJson(args["config"], args["json"])
	generator.run()
