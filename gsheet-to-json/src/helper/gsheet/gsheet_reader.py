#!/usr/bin/env python
'''
'''
from collections import defaultdict

import re
import importlib

import pygsheets

import urllib.request
from copy import deepcopy

from helper.logger import *
from helper.gsheet.gsheet_util import *

COLUMNS = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
			'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
			'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']

MASTER_TOC_COLUMNS = {
  "section" : {"availability": "must"},
  "heading" : {"availability": "must"},
  "process" : {"availability": "must"},
  "level" : {"availability": "must"},
  "content-type" : {"availability": "must"},
  "link" : {"availability": "must"},
  "break" : {"availability": "must"},
  "page-spec" : {"availability": "must"},
  "margin-spec" : {"availability": "must"},

  "landscape" : {"availability": "preferred", "value-if-missing": ""},
  "heading-style" : {"availability": "preferred", "value-if-missing": ""},
  "bookmark" : {"availability": "preferred", "value-if-missing": ""},

  "jpeg-quality" : {"availability": "preferred", "value-if-missing": '90'},
  "page-list" : {"availability": "preferred", "value-if-missing": ""},
  "autocrop" : {"availability": "preferred", "value-if-missing": False},
  "page-bg" : {"availability": "preferred", "value-if-missing": False},
 
  "hide-heading" : {"availability": "preferred", "value-if-missing": False},
  "header-first" : {"availability": "preferred", "value-if-missing": ""},
  "header-odd" : {"availability": "preferred", "value-if-missing": ""},
  "header-even" : {"availability": "preferred", "value-if-missing": ""},
  "footer-first" : {"availability": "preferred", "value-if-missing": ""},
  "footer-odd" : {"availability": "preferred", "value-if-missing": ""},
  "footer-even" : {"availability": "preferred", "value-if-missing": ""},
  "override-header" : {"availability": "preferred", "value-if-missing": False},
  "override-footer" : {"availability": "preferred", "value-if-missing": False},
  "background-image" : {"availability": "preferred", "value-if-missing": ""},
  "responsible" : {"availability": "preferred", "value-if-missing": ""},
  "reviewer" : {"availability": "preferred", "value-if-missing": ""},
  "status" : {"availability": "preferred", "value-if-missing": ""},
  "comment" : {"availability": "preferred", "value-if-missing": ""},
}

def process_gsheet(context, gsheet, parent, current_document_index, nesting_level):
    data = {'sections': []}

    # worksheet-cache is nested dictionary of gsheet->worksheet as two different sheets may have worksheets of same name
    # so keying by only worksheet name is not feasible
    if 'worksheet-cache' not in context:
        context['worksheet-cache'] = {}

    if gsheet.title not in context['worksheet-cache']:
        context['worksheet-cache'][gsheet.title] = {}


    # locate the index worksheet, it is a list, we take the first available from the list
    # TODO: for now it can also be a single worksheet for backword compatibility
    index_ws = None
    if isinstance(context['index-worksheet'], list):
        for ws_title in context['index-worksheet']:
            try:
                index_ws = gsheet.worksheet('title', ws_title)
                if index_ws:
                    break

            except Exception as e:
                warn(f"index worksheet [{ws_title}] not found", nesting_level=nesting_level)
                pass

        if index_ws is None:
            error(f"index worksheet not found from the list [{context['index-worksheet']}]")
            return

    else:
        ws_title = context['index-worksheet']
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
            trace(f"[{header:<20}] found at column [{COLUMNS[header_column_index]:>2}]", nesting_level=nesting_level)
            TOC_COLUMNS[header]['column'] = header_column_index

        else:
            trace(f"[{header:<20}] found at column [{COLUMNS[header_column_index]:>2}]. This column is not necessary for processing, will be ignored", nesting_level=nesting_level)

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

    toc_list = [toc for toc in toc_list[1:] if toc[TOC_COLUMNS['process'].get('column')] == 'Yes' and toc[TOC_COLUMNS['level'].get('column')] in [0, 1, 2, 3, 4, 5, 6]]

    section_index = 0
    for toc in toc_list:
        data['sections'].append(process_section(context=context, gsheet=gsheet, toc=toc, current_document_index=current_document_index, section_index=section_index, parent=parent, TOC_COLUMNS=TOC_COLUMNS, nesting_level=nesting_level))
        section_index = section_index + 1

    return data


def process_section(context, gsheet, toc, current_document_index, section_index, parent, TOC_COLUMNS, nesting_level):
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
        document_nesting_level = parent['section-meta']['nesting-level'] + 1
    else:
        document_nesting_level = 0

    d = {
        'section-meta' : {
            'document-name'         : gsheet.title,
            'document-index'        : current_document_index,
            'section-name'          : worksheet_name,
            'section-index'         : section_index,
            'nesting-level'         : document_nesting_level,
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
			'bookmark'           	: {translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='section', nesting_level=nesting_level+1): f"{str(translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='section', nesting_level=nesting_level+1))} {translate_dict_to_value(data_list=toc, dict_obj=TOC_COLUMNS, first_key='heading', nesting_level=nesting_level+1)}".strip()},

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
                d['header-first'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
                different_header_first_page = True

            else:
                d['header-first'] = None


            if d['header-odd'] != '' and d['header-odd'] is not None:
                new_section_data = {'section-prop': {'link': d['header-odd']}}
                d['header-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

            else:
                d['header-odd'] = None


            if d['header-even'] != '' and d['header-even'] is not None:
                new_section_data = {'section-prop': {'link': d['header-even']}}
                d['header-even'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
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
                d['footer-first'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
                different_footer_first_page = True

            else:
                d['footer-first'] = None


            if d['footer-odd'] != '' and d['footer-odd'] is not None:
                new_section_data = {'section-prop': {'link': d['footer-odd']}}
                d['footer-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

            else:
                d['footer-odd'] = None


            if d['footer-even'] != '' and d['footer-even'] is not None:
                new_section_data = {'section-prop': {'link': d['footer-even']}}
                d['footer-even'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
                different_footer_odd_even_pages = True

            else:
                d['footer-even'] = d['footer-odd']

    else:
        # process header, it may be text or link
        if d['header-first'] != '' and d['header-first'] is not None:
            new_section_data = {'section-prop': {'link': d['header-first']}}
            d['header-first'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
            different_header_first_page = True

        else:
            d['header-first'] = None


        if d['header-odd'] != '' and d['header-odd'] is not None:
            new_section_data = {'section-prop': {'link': d['header-odd']}}
            d['header-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

        else:
            d['header-odd'] = None


        if d['header-even'] != '' and d['header-even'] is not None:
            new_section_data = {'section-prop': {'link': d['header-even']}}
            d['header-even'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
            different_header_odd_even_pages = True

        else:
            d['header-even'] = d['header-odd']


        if d['footer-first'] != '' and d['footer-first'] is not None:
            new_section_data = {'section-prop': {'link': d['footer-first']}}
            d['footer-first'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
            different_footer_first_page = True

        else:
            d['footer-first'] = None


        if d['footer-odd'] != '' and d['footer-odd'] is not None:
            new_section_data = {'section-prop': {'link': d['footer-odd']}}
            d['footer-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

        else:
            d['footer-odd'] = None


        if d['footer-even'] != '' and d['footer-even'] is not None:
            new_section_data = {'section-prop': {'link': d['footer-even']}}
            d['footer-even'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
            different_footer_odd_even_pages = True

        else:
            d['footer-even'] = d['footer-odd']


    section_meta['different-firstpage'] = different_header_first_page or different_footer_first_page
    section_meta['different-odd-even-pages'] = different_header_odd_even_pages or different_footer_odd_even_pages

    # process 'background-image'
    if section_prop['background-image'] != '':
        bg_dict = download_image(url=section_prop['background-image'], tmp_dir=context['tmp-dir'], drive=context['drive'], nesting_level=nesting_level)
        if bg_dict:
            section_prop['background-image'] = bg_dict['file-path']
        else:
            section_prop['background-image'] = ''

    # import and use the specific processor
    if section_prop['link'] == '' or section_prop['link'] is None:
        d['contents'] = None

    else:
        module = importlib.import_module(f"processor.{section_prop['content-type']}_processor")
        d['contents'] = module.process(gsheet=gsheet, section_data=d, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

    return d
