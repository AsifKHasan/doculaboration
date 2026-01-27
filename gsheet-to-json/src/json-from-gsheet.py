#!/usr/bin/env python
'''
'''
import json
import time
import argparse

from ggle.google_services import GoogleServices
from ggle.gsheet_helper import GsheetHelper
from helper.config_service import ConfigService
from helper.logger import *


class JsonFromGsheet(object):

    def __init__(self):
        self.start_time = int(round(time.time() * 1000))

    def run(self, config_file, gsheet=None):
        # configuration
        config_service = ConfigService(config_file=config_file, nesting_level=0)

        # initialize GoogleServices
        google_services = GoogleServices(json_path=config_service._google_cred_json_path, nesting_level=0)


        # if gsheet name was provided as parameter, override the gsheet_list
        if gsheet:
            config_service._gsheet_list = [gsheet]

        # process gsheets one by one
        gsheet_helper = GsheetHelper()
        for gsheet_title in config_service._gsheet_list:
            output_json_path = f"{config_service._output_dir}/{gsheet_title}.json"
            gsheet_data = gsheet_helper.read_gsheet(gsheet_title=gsheet_title, gsheet_url=None, parent=None, nesting_level=0)
            with open(output_json_path, "w") as f:
                f.write(json.dumps(gsheet_data, sort_keys=False, indent=4))


        # tear down 
        self.end_time = int(round(time.time() * 1000))
        info(f"{gsheet_helper.current_document_index+1} documents/gsheets processed")
        info(f"script took {(self.end_time - self.start_time)/1000} seconds")
        # input("Press Enter to continue...")


if __name__ == '__main__':
	# construct the argument parse and parse the arguments
	ap = argparse.ArgumentParser()
	ap.add_argument("-c", "--config", required=True, help="configuration yml path")
	ap.add_argument("-g", "--gsheet", required=False, help="gsheet name to override gsheet list provided in configuration")
	args = vars(ap.parse_args())

	generator = JsonFromGsheet()
	generator.run(config_file=args["config"], gsheet=args["gsheet"])
