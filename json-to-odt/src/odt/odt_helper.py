#!/usr/bin/env python

''' Openoffice odt wrapper objects
'''

import time

from odf import opendocument

from helper.config_service import ConfigService
from helper.logger import *
from odt.odt_util import *

class OdtHelper(object):

    ''' constructor
    '''
    def __init__(self):
        self._odt = opendocument.load(ConfigService()._odt_template)


    ''' generate and save the odt
    '''
    def generate_and_save(self, section_list, nesting_level=0):
        self.start_time = int(round(time.time() * 1000))
        info(msg=f"generating odt ..", nesting_level=nesting_level)

        # font specs
        trace(f"registering fonts from conf/font-spec.yml", nesting_level=nesting_level+1)
        for k, v in ConfigService()._font_specs.items():
            if k != 'default':
                register_font(odt=self._odt, font_name=k, font_spec=v, nesting_level=nesting_level+2)

        # override styles
        trace(f"processing custom styles from conf/style-specs.yml", nesting_level=nesting_level+1)
        for k, v in ConfigService()._style_specs.items():
            update_style(odt=self._odt, style_key=k, style_spec=v, custom_styles=ConfigService()._style_specs, nesting_level=nesting_level+2)
        
        # process the sections
        section_list_to_odt(odt=self._odt, section_list=section_list, nesting_level=nesting_level+1)

        self.end_time = int(round(time.time() * 1000))
        info(msg=f"generated  odt .. {(self.end_time - self.start_time)/1000} seconds", nesting_level=nesting_level)

        # save the odt document
        output_odt_path = ConfigService()._output_odt_path
        info(msg=f"saving odt .. {output_odt_path}", nesting_level=nesting_level)
        self.start_time = int(round(time.time() * 1000))
        self._odt.save(output_odt_path)
        self.end_time = int(round(time.time() * 1000))
        info(msg=f"saved  odt .. {(self.end_time - self.start_time)/1000} seconds", nesting_level=nesting_level)

        # update indexes
        info(msg=f"updating index .. {output_odt_path}", nesting_level=nesting_level)
        self.start_time = int(round(time.time() * 1000))
        update_indexes(self._odt, output_odt_path, nesting_level=nesting_level+1)
        self.end_time = int(round(time.time() * 1000))
        info(msg=f"updated  index .. {(self.end_time - self.start_time)/1000} seconds", nesting_level=nesting_level)
