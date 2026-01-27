#!/usr/bin/env python
'''
'''
import time
import json
import argparse

from ggle.google_services import GoogleServices
from helper.config_service import ConfigService
from helper.logger import *
from odt.odt_helper import OdtHelper
from odt.odt_util import *

class OdtFromJson(object):

	def __init__(self):
		self.start_time = int(round(time.time() * 1000))


	def run(self, config_file, json_file=None, nesting_level=0):
		# configuration
		config_service = ConfigService(config_file=config_file, nesting_level=nesting_level)

		# initialize GoogleServices
		google_services = GoogleServices(json_path=config_service._google_cred_json_path, nesting_level=nesting_level)

		# if json name was provided as parameter, override the gsheet_list
		if json_file:
			config_service._json_list = [json_file]


		# process jsons one by one
		for json_file in config_service._json_list:
			config_service._input_json = config_service._output_dir / f"{json_file}.json"
			config_service._output_odt = config_service._output_dir / f"{json_file}.odt"
			with open(config_service._input_json, "r") as f:
				self._data = json.load(f)


			# odt-helper
			odt_helper = OdtHelper()
			odt_helper.generate_and_save(self._data['sections'])

			if config_service._generate_pdf:
				info(msg=f"generating pdf ..", nesting_level=nesting_level+1)
				pdf_start_time = int(round(time.time() * 1000))
				generate_pdf(self._CONFIG['files']['output-odt'], self._CONFIG['dirs']['output-dir'], nesting_level=nesting_level+1)
				self.end_time = int(round(time.time() * 1000))
				info(msg=f"generating pdf .. done {(self.end_time - pdf_start_time)/1000} seconds", nesting_level=nesting_level)

            # tear down
			self.end_time = int(round(time.time() * 1000))
			debug(msg=f"script took {(self.end_time - self.start_time)/1000} seconds")


if __name__ == '__main__':
	# construct the argument parse and parse the arguments
	ap = argparse.ArgumentParser()
	ap.add_argument("-c", "--config", required=True, help="configuration yml path")
	ap.add_argument("-j", "--json", required=False, help="json name to override json list provided in configuration")
	args = vars(ap.parse_args())

	generator = OdtFromJson()
	generator.run(config_file=args["config"], json_file=args["json"], nesting_level=0)
