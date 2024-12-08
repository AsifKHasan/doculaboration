#!/usr/bin/env python3
'''
'''
from collections import defaultdict

import re
import importlib

import pygsheets
import pprint

import urllib.request

from helper.logger import *
from helper.gsheet.gsheet_util import *

COLUMNS = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
			'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
			'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']


TOC_COLUMNS = {
  "section" : {"availability": "must"},
  "heading" : {"availability": "must"},
  "process" : {"availability": "must"},
  "level" : {"availability": "must"},
  "content-type" : {"availability": "must"},
  "link" : {"availability": "must"},
  "break" : {"availability": "must"},
  "landscape" : {"availability": "must"},
  "page-spec" : {"availability": "must"},
  "margin-spec" : {"availability": "must"},
  "hide-pageno" : {"availability": "preferred"},
  "hide-heading" : {"availability": "preferred"},
  "different-firstpage" : {"availability": "preferred"},
  "header-first" : {"availability": "preferred"},
  "header-odd" : {"availability": "preferred"},
  "header-even" : {"availability": "preferred"},
  "footer-first" : {"availability": "preferred"},
  "footer-odd" : {"availability": "preferred"},
  "footer-even" : {"availability": "preferred"},
  "override-header" : {"availability": "preferred"},
  "override-footer" : {"availability": "preferred"},
  "background-image" : {"availability": "preferred"},
  "responsible" : {"availability": "preferred"},
  "reviewer" : {"availability": "preferred"},
  "status" : {"availability": "preferred"},
  "comment" : {"availability": "preferred"},
}


def process_gsheet(context, gsheet, parent, current_document_index, nesting_level):
    data = {'sections': []}

    # worksheet-cache is nested dictionary of gsheet->worksheet as two different sheets may have worksheets of same name
    # so keying by only worksheet name is not feasible
    if 'worksheet-cache' not in context:
        context['worksheet-cache'] = {}

    if gsheet.title not in context['worksheet-cache']:
        context['worksheet-cache'][gsheet.title] = {}

    ws_title = context['index-worksheet']
    ws = gsheet.worksheet('title', ws_title)

    # we expect toc data to start from row 2
    last_column = COLUMNS[ws.cols-1]
    toc_list = ws.get_values(start='A2', end=f"{last_column}{ws.rows}", returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=False, value_render='FORMULA')

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
            trace(f"header [{toc_column_key:<20}], which is [{toc_column_value['availability']}] found in column [{COLUMNS[toc_column_value['column']]:>2}]", nesting_level=nesting_level)

        else:
            if toc_column_value['availability'] == 'must':
                error(f"header [{toc_column_key:<20}], which is must, not found in any column. Exiting ...")
                failed = True

            elif toc_column_value['availability'] == 'preferred':
                warn(f"header [{toc_column_key:<20}], which is preferred, not found in any column. Consider adding the column", nesting_level=nesting_level)
        
    if failed:
        exit(-1)











    exit(0)
    toc_list = [toc for toc in toc_list if toc[2] == 'Yes' and toc[3] in [0, 1, 2, 3, 4, 5, 6]]

    section_index = 0
    for toc in toc_list:
        data['sections'].append(process_section(context=context, gsheet=gsheet, toc=toc, current_document_index=current_document_index, section_index=section_index, parent=parent, nesting_level=nesting_level))
        section_index = section_index + 1

    return data


def process_section(context, gsheet, toc, current_document_index, section_index, parent, nesting_level):
    # TODO: some columns may have formula, parse those
    # link column (F, toc[5] may be a formula), parse it
    if toc[4] in ['gsheet', 'pdf']:
        link_name, link_target = get_gsheet_link(toc[5], nesting_level=nesting_level)
        worksheet_name = link_name

    elif toc[4] == 'table':
        link_name, link_target = get_worksheet_link(toc[5], nesting_level=nesting_level), None
        worksheet_name = link_name

    else:
        link_name, link_target = toc[5], None
        worksheet_name = toc[4]

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
            'label'                 : str(toc[0]),
            'heading'               : toc[1],
            'process'               : toc[2],
            'level'                 : int(toc[3]),
            'content-type'          : toc[4],
            'link'                  : link_name,
            'link-target'           : link_target,
            'page-break'            : True if toc[6] == "page" else False,
            'section-break'         : True if toc[6] == "section" else False,
            'landscape'             : True if toc[7] == "Yes" else False,
            'page-spec'             : toc[8],
            'margin-spec'           : toc[9],
            'hide-pageno'           : True if toc[10] == "Yes" else False,
            'hide-heading'          : True if toc[11] == "Yes" else False,
            'different-firstpage'   : True if toc[12] == "Yes" else False,
            'override-header'       : True if toc[19] == "Yes" else False,
            'override-footer'       : True if toc[20] == "Yes" else False,
            'background-image'      : toc[21].strip(),
            'responsible'           : toc[22].strip(),
            'reviewer'              : toc[23].strip(),
            'status'                : toc[24].strip(),
        },
        'header-odd'            : get_worksheet_link(toc[14]),
        'header-even'           : get_worksheet_link(toc[15]),
        'footer-odd'            : get_worksheet_link(toc[17]),
        'footer-even'           : get_worksheet_link(toc[18]),
        'header-first'          : get_worksheet_link(toc[13]),
        'footer-first'          : get_worksheet_link(toc[16]),
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

    different_odd_even_pages = False

    # the gsheet is a child gsheet, called from a parent gsheet, so header processing depends on override flags
    module = importlib.import_module('processor.table_processor')
    if parent:
        parent_section_meta = parent['section-meta']
        parent_section_prop = parent['section-prop']

        if parent_section_prop['override-header']:
            # debug(f".. this is a child gsheet : header OVERRIDDEN")
            section_prop['different-firstpage'] = parent_section_prop['different-firstpage']
            d['header-first'] = parent['header-first']
            d['header-odd'] = parent['header-odd']
            d['header-even'] = parent['header-even']

        else:
            # debug(f".. child gsheet's header is NOT overridden")
            if section_prop['different-firstpage']:
                if d['header-first'] != '' and d['header-first'] is not None:
                    new_section_data = {'section-prop': {'link': d['header-first']}}
                    d['header-first'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

                else:
                    d['header-first'] = None

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
                different_odd_even_pages = True

            else:
                d['header-even'] = d['header-odd']


        if parent_section_prop['override-footer']:
            # debug(f".. this is a child gsheet : footer OVERRIDDEN")
            section_prop['different-firstpage'] = parent_section_prop['different-firstpage']
            d['footer-first'] = parent['footer-first']
            d['footer-odd'] = parent['footer-odd']
            d['footer-even'] = parent['footer-even']

        else:
            # debug(f".. child gsheet's footer is NOT overridden")
            if section_prop['different-firstpage']:
                if d['footer-first'] != '' and d['footer-first'] is not None:
                    new_section_data = {'section-prop': {'link': d['footer-first']}}
                    d['footer-first'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

                else:
                    d['footer-first'] = None

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
                different_odd_even_pages = True

            else:
                d['footer-even'] = d['footer-odd']

    else:
        # process header, it may be text or link
        if section_prop['different-firstpage']:
            if d['header-first'] != '' and d['header-first'] is not None:
                new_section_data = {'section-prop': {'link': d['header-first']}}
                d['header-first'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

            else:
                d['header-first'] = None


            if d['footer-first'] != '' and d['footer-first'] is not None:
                new_section_data = {'section-prop': {'link': d['footer-first']}}
                d['footer-first'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

            else:
                d['footer-first'] = None

        else:
            d['header-first'] = None
            d['footer-first'] = None


        if d['header-odd'] != '' and d['header-odd'] is not None:
            new_section_data = {'section-prop': {'link': d['header-odd']}}
            d['header-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

        else:
            d['header-odd'] = None


        if d['header-even'] != '' and d['header-even'] is not None:
            new_section_data = {'section-prop': {'link': d['header-even']}}
            d['header-even'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
            different_odd_even_pages = True

        else:
            d['header-even'] = d['header-odd']


        if d['footer-odd'] != '' and d['footer-odd'] is not None:
            new_section_data = {'section-prop': {'link': d['footer-odd']}}
            d['footer-odd'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)

        else:
            d['footer-odd'] = None


        if d['footer-even'] != '' and d['footer-even'] is not None:
            new_section_data = {'section-prop': {'link': d['footer-even']}}
            d['footer-even'] = module.process(gsheet=gsheet, section_data=new_section_data, context=context, current_document_index=current_document_index, nesting_level=nesting_level)
            different_odd_even_pages = True

        else:
            d['footer-even'] = d['footer-odd']


    section_meta['different-odd-even-pages'] = different_odd_even_pages

    # process 'background-image'
    if section_prop['background-image'] != '':
        bg_dict = download_image(url=section_prop['background-image'], tmp_dir=context['tmp-dir'], nesting_level=nesting_level)
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
