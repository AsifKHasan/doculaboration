#!/usr/bin/env python3
'''
'''
import os
import sys
import json
import importlib
import time
import yaml
import datetime
import argparse
import pprint

from helper.logger import *
from helper.pandoc.pandoc_helper import PandocHelper
from helper.pandoc.pandoc_util import *

class PandocFromJson(object):

	def __init__(self, config_path):
		self.start_time = int(round(time.time() * 1000))
		self._config_path = os.path.abspath(config_path)
		self._data = {}

	def generate_pandoc(self):
		for section in self._data['sections']:
			content_type = section['content-type']

			# force table formatter for gsheet content
			if content_type == 'gsheet': content_type = 'table'

			module = importlib.import_module('formatter.{0}_formatter'.format(content_type))
			section_doc = module.generate(section, self._doc, self._pandochelper._sections, self._CONFIG)
			self._doc = self._doc + section_doc

		self._pandochelper.save(self._doc)

	def run(self):
		self.set_up()
		# process jsons one by one
		for json in self._CONFIG['jsons']:
			self._CONFIG['files']['input-json'] = os.path.abspath('{0}/{1}.json'.format(self._CONFIG['dirs']['output-dir'], json))
			self.load_json()

			# pandoc-helper
			self._CONFIG['files']['output-pandoc'] = os.path.abspath('{0}/{1}.md'.format(self._CONFIG['dirs']['output-dir'], json))
			self._pandochelper = PandocHelper(self._CONFIG['files']['pandoc-styles'], self._CONFIG['files']['output-pandoc'])
			self._doc = self._pandochelper.init()
			self.generate_pandoc()

			self.tear_down()

	def set_up(self):
		# configuration
		self._CONFIG = yaml.load(open(self._config_path, 'r', encoding='utf-8'), Loader=yaml.FullLoader)
		config_dir = os.path.dirname(self._config_path)

		self._CONFIG['dirs']['output-dir'] = os.path.abspath('{0}/{1}'.format(config_dir, self._CONFIG['dirs']['output-dir']))
		self._CONFIG['dirs']['temp-dir'] = os.path.abspath('{0}/tmp'.format(self._CONFIG['dirs']['output-dir']))
		if not os.path.exists(self._CONFIG['dirs']['temp-dir']):
			os.makedirs(self._CONFIG['dirs']['temp-dir'])

		self._CONFIG['files']['pandoc-styles'] = os.path.abspath('{0}/{1}'.format(config_dir, self._CONFIG['files']['pandoc-styles']))

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
	args = vars(ap.parse_args())

	generator = PandocFromJson(args["config"])
	generator.run()
