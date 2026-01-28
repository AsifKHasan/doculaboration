#!/usr/bin/env python
'''
'''

import sys
import pygsheets
import importlib
from copy import deepcopy

from ggle.google_services import GoogleServices
from helper.config_service import ConfigService
from helper.logger import *
from helper.util import *

class GsheetHelper(object):

    _instance = None

    ''' class constructor
    '''
    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(GsheetHelper, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance


    ''' initialize the helper
    '''
    def __init__(self, nesting_level=0):
        if self._initialized:
            return

        self.google_services = GoogleServices()
        self.config_service = ConfigService()

        self.google_services.pygsheet = pygsheets.authorize(custom_credentials=self.google_services._credential)
        self.worksheet_cache = {}
        self.gsheet_data = {}
        self.current_document_index = -1

        self._initialized = True


    ''' read the gsheet
    '''
    def read_gsheet(self, gsheet_title, gsheet_url=None, parent=None, nesting_level=0):
        gsheet = None
        for i in range(0, self.config_service._gsheet_read_try_count):
            try:
                if gsheet_url:
                    gsheet_id = gsheet_id_from_url(url=gsheet_url, nesting_level=nesting_level)
                    debug(f"opening gsheet id = {gsheet_id}", nesting_level=nesting_level)
                    gsheet = self.google_services.pygsheet.open_by_url(gsheet_url)
                    debug(f"opened  gsheet id = {gsheet_id}", nesting_level=nesting_level)
                else:
                    query = f'name = "{gsheet_title}"'
                    
                    debug(f"opening gsheet : [{gsheet_title}]", nesting_level=nesting_level)
                    gsheets = self.google_services.pygsheet.open_all(query=query)
                    for gsheet in gsheets:
                        # Call the API to get permissions
                        results = self.google_services.drive_api.permissions().list(fileId=gsheet.id, fields="permissions(id, emailAddress, role, displayName)").execute()
                        trace(f"[{gsheet_title}] found with id [{gsheet.id}]", nesting_level=nesting_level)
                        permissions = results.get('permissions', [])
                        for perm in permissions:
                            trace(f"{perm['role'].upper()}: {perm.get('displayName')} ({perm.get('emailAddress')})", nesting_level=nesting_level+1)

                    if len(gsheets) > 1:
                        error(f"[{len(gsheets)}] gsheets found with the name [{gsheet_title}] .. quiting", nesting_level=nesting_level)
                        sys.exit(1)

                    elif len(gsheets) == 1:
                        gsheet = gsheets[0]
                        
                    else:
                        error(f"no gsheet found with the name [{gsheet_title}] .. quiting", nesting_level=nesting_level)
                        sys.exit(1)

                    gsheet_id = gsheet.id
                    debug(f"opened  gsheet : [{gsheet_title}] [{gsheet_id}]", nesting_level=nesting_level)

                # optimization - read the full gsheet
                debug(f"reading gsheet : [{gsheet_title}] [{gsheet_id}]", nesting_level=nesting_level)

                response = get_gsheet_data(sheets_service=self.google_services.sheets_api, spreadsheet_id=gsheet.id, nesting_level=nesting_level+1)

                # make a dictionary key'ed by worksheet_name
                response = {sheet['properties']['title']: sheet for sheet in response['sheets']}
                self.gsheet_data[gsheet_title] = response

                debug(f"read    gsheet : [{gsheet_title}] [{gsheet_id}]", nesting_level=nesting_level)

                break

            except Exception as err:
                print(err)
                warn(f"gsheet {gsheet_title} read request (attempt {i}) failed, waiting for {self.config_service._gsheet_read_wait_seconds} seconds before trying again", nesting_level=nesting_level)
                time.sleep(float(self.config_service._gsheet_read_wait_seconds))

        if gsheet is None:
            error('gsheet read request failed, quiting', nesting_level=nesting_level)
            sys.exit(1)

        self.current_document_index = self.current_document_index + 1
        data = self.process_gsheet(gsheet=gsheet, parent=parent, current_document_index=self.current_document_index, nesting_level=nesting_level+1)

        return data

    ''' process gsheet from the toc
        worksheet_cache: nested dictionary of gsheet->worksheet as two different sheets may have worksheets of same name, so keying by only worksheet name is not feasible
    '''
    def process_gsheet(self, gsheet, parent, current_document_index, nesting_level=0):
        data = {'sections': []}

        # worksheet_cache: nested dictionary of gsheet->worksheet as two different sheets may have worksheets of same name, so keying by only worksheet name is not feasible
        if gsheet.title not in self.worksheet_cache:
            self.worksheet_cache[gsheet.title] = {}


        # locate the index worksheet, it is a list, we take the first available from the list
        # TODO: for now it can also be a single worksheet for backword compatibility
        index_ws = None
        if isinstance(ConfigService()._index_worksheet, list):
            for ws_title in ConfigService()._index_worksheet:
                try:
                    index_ws = gsheet.worksheet('title', ws_title)
                    if index_ws:
                        break

                except Exception as e:
                    warn(f"index worksheet [{ws_title}] not found", nesting_level=nesting_level)
                    pass

            if index_ws is None:
                error(f"index worksheet not found from the list [{ConfigService()._index_worksheet}]")
                return

        else:
            ws_title = ConfigService()._index_worksheet
            try:
                index_ws = gsheet.worksheet('title', ws_title)
            except Exception as e:
                error(f"index worksheet [{ws_title}] not found")
                return

        if index_ws:
            debug(f"index worksheet [{ws_title}] found", nesting_level=nesting_level)

        # we expect toc data to start from row 2
        last_column = COLUMNS[index_ws.cols-1]
        toc_list = index_ws.get_values(start='A2', end=f"{last_column}{index_ws.rows}", returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=False, value_render='FORMULA')

        # make a deep copy of MASTER_TOC_COLUMNS for this gsheet
        TOC_COLUMNS = deepcopy(MASTER_TOC_COLUMNS)

        # the first item is the heading for all columns which says what data is in which column
        header_list = toc_list[0]
        header_column_index = 0
        for header in header_list:
            # see if it is a header we know of
            toc_column = TOC_COLUMNS.get(header, None)
            if toc_column:
                # trace(f"[{header:<20}] found at column [{COLUMNS[header_column_index]:>2}]", nesting_level=nesting_level)
                TOC_COLUMNS[header]['column'] = header_column_index

            else:
                warn(f"[{header:<20}] found at column [{COLUMNS[header_column_index]:>2}]. This column is not necessary for processing, will be ignored", nesting_level=nesting_level)

            header_column_index = header_column_index + 1

        # check if we have any must column missing
        failed = False
        for toc_column_key, toc_column_value in TOC_COLUMNS.items():
            if 'column' in toc_column_value:
                trace(f"header [{toc_column_key:<20}], which is [{toc_column_value['availability']}] found in column [{COLUMNS[toc_column_value.get('column')]:>2}]", nesting_level=nesting_level)

            else:
                if toc_column_value['availability'] == 'must':
                    error(f"header [{toc_column_key:<20}], which is must, not found in any column. Exiting ...")
                    failed = True

                elif toc_column_value['availability'] == 'preferred':
                    warn(f"header [{toc_column_key:<20}], which is preferred, not found in any column. Consider adding the column", nesting_level=nesting_level)

        if failed:
            exit(-1)

        # toc_list = [toc for toc in toc_list[1:] if toc[TOC_COLUMNS['process'].get('column')] == 'Yes']
        toc_list_to_process = []
        for row, toc in enumerate(toc_list[1:], 3):
            if toc[TOC_COLUMNS['process'].get('column')] == 'Yes':
                # check if any of the columns that do not allow blanks is missing values
                for k, v in TOC_COLUMNS.items():
                    if 'column' in v and 'blank-allowed' in v:
                        if v['blank-allowed'] == False:
                            if toc[v['column']] == '':
                                error(f"toc row [{row}] must have some value for [{k}] in column [{COLUMNS[v['column']]}].. exiting", nesting_level=nesting_level)
                                exit(1)

                # trace(f"toc row [{row}] will be processed", nesting_level=nesting_level)
                toc_list_to_process.append(toc)
                # toc[TOC_COLUMNS['level'].get('column')] in [0, 1, 2, 3, 4, 5, 6]

            else:
                # trace(f"toc row [{row}] will NOT be processed", nesting_level=nesting_level)
                pass


        section_index = 0
        for section_index, toc in enumerate(toc_list_to_process):
            data['sections'].append(self.process_section(gsheet=gsheet, toc=toc, current_document_index=current_document_index, section_index=section_index, parent=parent, TOC_COLUMNS=TOC_COLUMNS, nesting_level=nesting_level))

        return data


    ''' process a toc section
    '''
    def process_section(self, gsheet, toc, current_document_index, section_index, parent, TOC_COLUMNS, nesting_level):
        # TODO: some columns may have formula, parse those
        # link column (link) may be a formula, parse it
        if toc[TOC_COLUMNS['content-type'].get('column')] in ['gsheet', 'pdf']:
            link_name, link_target = get_gsheet_link(toc[TOC_COLUMNS['link'].get('column')], nesting_level=nesting_level)
            worksheet_name = link_name

        elif toc[TOC_COLUMNS['content-type'].get('column')] == 'table':
            link_name, link_target = get_worksheet_link(toc[TOC_COLUMNS['link'].get('column')], nesting_level=nesting_level), None
            worksheet_name = link_name

        else:
            link_name, link_target = toc[TOC_COLUMNS['link'].get('column')], None
            worksheet_name = toc[TOC_COLUMNS['content-type'].get('column')]

        if section_index == 0:
            document_first_section = True
            if current_document_index == 0:
                first_section = True

            else:
                first_section = False

        else:
            document_first_section = False
            first_section = False


        # transform to a dict
        if parent:
            document_nesting_depth = parent['section-meta']['document-nesting-depth'] + 1
        else:
            document_nesting_depth = 0

        d = {
            'section-meta' : {
                'document-name'         : gsheet.title,
                'document-index'        : current_document_index,
                'section-name'          : worksheet_name,
                'section-index'         : section_index,
                'document-nesting-depth': document_nesting_depth,
                'first-section'         : first_section,
                'document-first-section': document_first_section,
            },
            'section-prop' : {
                'label'                 : str(toc[TOC_COLUMNS['section'].get('column')]),
                'heading'               : toc[TOC_COLUMNS['heading'].get('column')],
                'process'               : toc[TOC_COLUMNS['process'].get('column')],
                'level'                 : int(toc[TOC_COLUMNS['level'].get('column')]),
                'content-type'          : toc[TOC_COLUMNS['content-type'].get('column')],
                'link'                  : link_name,
                'link-target'           : link_target,
                'page-break'            : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='break', look_up_value='page', nesting_level=nesting_level+1),
                'section-break'         : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='break', look_up_value='section', nesting_level=nesting_level+1),
                'page-spec'             : toc[TOC_COLUMNS['page-spec'].get('column')],
                'margin-spec'           : toc[TOC_COLUMNS['margin-spec'].get('column')],

                'landscape'             : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='landscape', look_up_value='Yes', nesting_level=nesting_level+1),
                'heading-style'         : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='heading-style', nesting_level=nesting_level+1),
                'bookmark'           	: {translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='bookmark', nesting_level=nesting_level+1): f"{str(translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='section', nesting_level=nesting_level+1))} {translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='heading', nesting_level=nesting_level+1)}".strip()},

                'jpeg-quality'          : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='jpeg-quality', nesting_level=nesting_level+1),
                'page-list'             : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='page-list', nesting_level=nesting_level+1),
                'autocrop'              : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='autocrop', look_up_value='Yes', nesting_level=nesting_level+1),
                'page-bg'               : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='page-bg', look_up_value='Yes', nesting_level=nesting_level+1),

                'hide-heading'          : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='hide-heading', look_up_value='Yes', nesting_level=nesting_level+1),
                'override-header'       : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='override-header', look_up_value='Yes', nesting_level=nesting_level+1),
                'override-footer'       : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='override-footer', look_up_value='Yes', nesting_level=nesting_level+1),
                'background-image'      : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='background-image', nesting_level=nesting_level+1),

                'responsible'           : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='responsible', nesting_level=nesting_level+1),
                'reviewer'              : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='reviewer', nesting_level=nesting_level+1),
                'status'                : translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='status', nesting_level=nesting_level+1),
            },
            'header-odd'            : get_worksheet_link(translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='header-odd', nesting_level=nesting_level+1)),
            'header-even'           : get_worksheet_link(translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='header-even', nesting_level=nesting_level+1)),
            'header-first'          : get_worksheet_link(translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='header-first', nesting_level=nesting_level+1)),
            'footer-odd'            : get_worksheet_link(translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='footer-odd', nesting_level=nesting_level+1)),
            'footer-even'           : get_worksheet_link(translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='footer-even', nesting_level=nesting_level+1)),
            'footer-first'          : get_worksheet_link(translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='footer-first', nesting_level=nesting_level+1)),
            }

        section_meta = d['section-meta']
        section_prop = d['section-prop']

        # the page-spec, margin-spec is overridden by parent page-spec
        if parent:
            section_prop['page-spec'] = parent['section-prop']['page-spec']
            section_prop['margin-spec'] = parent['section-prop']['margin-spec']

        # derived keys
        if section_prop['landscape']:
            section_meta['orientation'] = 'landscape'
        else:
            section_meta['orientation'] = 'portrait'

        section_meta['page-layout'] = f"{section_prop['page-spec']}-{section_meta['orientation']}-{section_prop['margin-spec']}"

        different_header_first_page = False
        different_footer_first_page = False

        different_header_odd_even_pages = False
        different_footer_odd_even_pages = False

        # the gsheet is a child gsheet, called from a parent gsheet, so header processing depends on override flags
        module = importlib.import_module('processor.table_processor')
        if parent:
            parent_section_meta = parent['section-meta']
            parent_section_prop = parent['section-prop']

            if parent_section_prop['override-header']:
                # debug(f".. this is a child gsheet : header OVERRIDDEN")
                section_meta['different-firstpage'] = parent_section_meta['different-firstpage']
                d['header-first'] = parent['header-first']
                d['header-odd'] = parent['header-odd']
                d['header-even'] = parent['header-even']

            else:
                # debug(f".. child gsheet's header is NOT overridden")
                if d['header-first'] != '' and d['header-first'] is not None:
                    new_section_data = {'section-prop': {'link': d['header-first']}}
                    d['header-first'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)
                    different_header_first_page = True

                else:
                    d['header-first'] = None


                if d['header-odd'] != '' and d['header-odd'] is not None:
                    new_section_data = {'section-prop': {'link': d['header-odd']}}
                    d['header-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)

                else:
                    d['header-odd'] = None


                if d['header-even'] != '' and d['header-even'] is not None:
                    new_section_data = {'section-prop': {'link': d['header-even']}}
                    d['header-even'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)
                    different_header_odd_even_pages = True

                else:
                    d['header-even'] = d['header-odd']


            if parent_section_prop['override-footer']:
                # debug(f".. this is a child gsheet : footer OVERRIDDEN")
                section_meta['different-firstpage'] = parent_section_meta['different-firstpage']
                d['footer-first'] = parent['footer-first']
                d['footer-odd'] = parent['footer-odd']
                d['footer-even'] = parent['footer-even']

            else:
                # debug(f".. child gsheet's footer is NOT overridden")
                if d['footer-first'] != '' and d['footer-first'] is not None:
                    new_section_data = {'section-prop': {'link': d['footer-first']}}
                    d['footer-first'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)
                    different_footer_first_page = True

                else:
                    d['footer-first'] = None


                if d['footer-odd'] != '' and d['footer-odd'] is not None:
                    new_section_data = {'section-prop': {'link': d['footer-odd']}}
                    d['footer-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)

                else:
                    d['footer-odd'] = None


                if d['footer-even'] != '' and d['footer-even'] is not None:
                    new_section_data = {'section-prop': {'link': d['footer-even']}}
                    d['footer-even'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)
                    different_footer_odd_even_pages = True

                else:
                    d['footer-even'] = d['footer-odd']

        else:
            # process header, it may be text or link
            if d['header-first'] != '' and d['header-first'] is not None:
                new_section_data = {'section-prop': {'link': d['header-first']}}
                d['header-first'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)
                different_header_first_page = True

            else:
                d['header-first'] = None


            if d['header-odd'] != '' and d['header-odd'] is not None:
                new_section_data = {'section-prop': {'link': d['header-odd']}}
                d['header-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)

            else:
                d['header-odd'] = None


            if d['header-even'] != '' and d['header-even'] is not None:
                new_section_data = {'section-prop': {'link': d['header-even']}}
                d['header-even'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)
                different_header_odd_even_pages = True

            else:
                d['header-even'] = d['header-odd']


            if d['footer-first'] != '' and d['footer-first'] is not None:
                new_section_data = {'section-prop': {'link': d['footer-first']}}
                d['footer-first'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)
                different_footer_first_page = True

            else:
                d['footer-first'] = None


            if d['footer-odd'] != '' and d['footer-odd'] is not None:
                new_section_data = {'section-prop': {'link': d['footer-odd']}}
                d['footer-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)

            else:
                d['footer-odd'] = None


            if d['footer-even'] != '' and d['footer-even'] is not None:
                new_section_data = {'section-prop': {'link': d['footer-even']}}
                d['footer-even'] = module.process(gsheet=gsheet, section_data=new_section_data, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)
                different_footer_odd_even_pages = True

            else:
                d['footer-even'] = d['footer-odd']


        section_meta['different-firstpage'] = different_header_first_page or different_footer_first_page
        section_meta['different-odd-even-pages'] = different_header_odd_even_pages or different_footer_odd_even_pages

        # process 'background-image'
        if section_prop['background-image'] != '':
            bg_dict = download_image(drive_service=GoogleServices().drive_api, url=section_prop['background-image'], title=None, tmp_dir=ConfigService()._temp_dir, nesting_level=nesting_level)
            if bg_dict:
                section_prop['background-image'] = bg_dict['file-path']
            else:
                section_prop['background-image'] = ''

        # import and use the specific processor
        if section_prop['link'] == '' or section_prop['link'] is None:
            d['contents'] = None

        else:
            module = importlib.import_module(f"processor.{section_prop['content-type']}_processor")
            d['contents'] = module.process(gsheet=gsheet, section_data=d, worksheet_cache=self.worksheet_cache, gsheet_data=self.gsheet_data, current_document_index=current_document_index, nesting_level=nesting_level)

        return d
