#!/usr/bin/env python
'''
'''
from collections import defaultdict

import re
import importlib

import pygsheets

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
  "page-spec" : {"availability": "must"},
  "margin-spec" : {"availability": "must"},
  "landscape" : {"availability": "must"},
  "bookmark" : {"availability": "preferred"},
  "autocrop" : {"availability": "preferred"},
  "page-bg" : {"availability": "preferred"},
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

section_column = None
heading_column = None
process_column = None
level_column = None
content_type_column = None
link_column = None
break_column = None
page_spec_column = None
margin_spec_column = None
landscape_column = None
bookmark_column = None
autocrop_column = None
page_bg_column = None
hide_pageno_column = None
hide_heading_column = None
different_firstpage_column = None
header_first_column = None
header_odd_column = None
header_even_column = None
footer_first_column = None
footer_odd_column = None
footer_even_column = None
override_header_column = None
override_footer_column = None
background_image_column = None
responsible_column = None
reviewer_column = None
status_column = None
comment_column = None

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

    global section_column, heading_column, process_column, level_column, content_type_column, link_column, break_column, page_spec_column, margin_spec_column
    global landscape_column, bookmark_column, autocrop_column, page_bg_column, hide_pageno_column, hide_heading_column
    global different_firstpage_column, header_first_column, header_odd_column, header_even_column, footer_first_column, footer_odd_column, footer_even_column
    global override_header_column, override_footer_column, background_image_column
    global responsible_column, reviewer_column,status_column, comment_column

    section_column = TOC_COLUMNS['section']['column'] if 'section' in TOC_COLUMNS and 'column' in TOC_COLUMNS['section'] else None
    heading_column = TOC_COLUMNS['heading']['column'] if 'heading' in TOC_COLUMNS and 'column' in TOC_COLUMNS['heading'] else None
    process_column = TOC_COLUMNS['process']['column'] if 'process' in TOC_COLUMNS and 'column' in TOC_COLUMNS['process'] else None
    level_column = TOC_COLUMNS['level']['column'] if 'level' in TOC_COLUMNS and 'column' in TOC_COLUMNS['level'] else None
    content_type_column = TOC_COLUMNS['content-type']['column'] if 'content-type' in TOC_COLUMNS and 'column' in TOC_COLUMNS['content-type'] else None
    link_column = TOC_COLUMNS['link']['column'] if 'link' in TOC_COLUMNS and 'column' in TOC_COLUMNS['link'] else None
    break_column = TOC_COLUMNS['break']['column'] if 'break' in TOC_COLUMNS and 'column' in TOC_COLUMNS['break'] else None
    page_spec_column = TOC_COLUMNS['page-spec']['column'] if 'page-spec' in TOC_COLUMNS and 'column' in TOC_COLUMNS['page-spec'] else None
    margin_spec_column = TOC_COLUMNS['margin-spec']['column'] if 'margin-spec' in TOC_COLUMNS and 'column' in TOC_COLUMNS['margin-spec'] else None
    landscape_column = TOC_COLUMNS['landscape']['column'] if 'landscape' in TOC_COLUMNS and 'column' in TOC_COLUMNS['landscape'] else None
    bookmark_column = TOC_COLUMNS['bookmark']['column'] if 'bookmark' in TOC_COLUMNS and 'column' in TOC_COLUMNS['bookmark'] else None
    autocrop_column = TOC_COLUMNS['autocrop']['column'] if 'autocrop' in TOC_COLUMNS and 'column' in TOC_COLUMNS['autocrop'] else None
    page_bg_column = TOC_COLUMNS['page-bg']['column'] if 'page-bg' in TOC_COLUMNS and 'column' in TOC_COLUMNS['page-bg'] else None
    hide_pageno_column = TOC_COLUMNS['hide-pageno']['column'] if 'hide-pageno' in TOC_COLUMNS and 'column' in TOC_COLUMNS['hide-pageno'] else None
    hide_heading_column = TOC_COLUMNS['hide-heading']['column'] if 'hide-heading' in TOC_COLUMNS and 'column' in TOC_COLUMNS['hide-heading'] else None
    different_firstpage_column = TOC_COLUMNS['different-firstpage']['column'] if 'different-firstpage' in TOC_COLUMNS and 'column' in TOC_COLUMNS['different-firstpage'] else None
    header_first_column = TOC_COLUMNS['header-first']['column'] if 'header-first' in TOC_COLUMNS and 'column' in TOC_COLUMNS['header-first'] else None
    header_odd_column = TOC_COLUMNS['header-odd']['column'] if 'header-odd' in TOC_COLUMNS and 'column' in TOC_COLUMNS['header-odd'] else None
    header_even_column = TOC_COLUMNS['header-even']['column'] if 'header-even' in TOC_COLUMNS and 'column' in TOC_COLUMNS['header-even'] else None
    footer_first_column = TOC_COLUMNS['footer-first']['column'] if 'footer-first' in TOC_COLUMNS and 'column' in TOC_COLUMNS['footer-first'] else None
    footer_odd_column = TOC_COLUMNS['footer-odd']['column'] if 'footer-odd' in TOC_COLUMNS and 'column' in TOC_COLUMNS['footer-odd'] else None
    footer_even_column = TOC_COLUMNS['footer-even']['column'] if 'footer-even' in TOC_COLUMNS and 'column' in TOC_COLUMNS['footer-even'] else None
    override_header_column = TOC_COLUMNS['override-header']['column'] if 'override-header' in TOC_COLUMNS and 'column' in TOC_COLUMNS['override-header'] else None
    override_footer_column = TOC_COLUMNS['override-footer']['column'] if 'override-footer' in TOC_COLUMNS and 'column' in TOC_COLUMNS['override-footer'] else None
    background_image_column = TOC_COLUMNS['background-image']['column'] if 'background-image' in TOC_COLUMNS and 'column' in TOC_COLUMNS['background-image'] else None
    responsible_column = TOC_COLUMNS['responsible']['column'] if 'responsible' in TOC_COLUMNS and 'column' in TOC_COLUMNS['responsible'] else None
    reviewer_column = TOC_COLUMNS['reviewer']['column'] if 'reviewer' in TOC_COLUMNS and 'column' in TOC_COLUMNS['reviewer'] else None
    status_column = TOC_COLUMNS['status']['column'] if 'status' in TOC_COLUMNS and 'column' in TOC_COLUMNS['status'] else None
    comment_column = TOC_COLUMNS['comment']['column'] if 'comment' in TOC_COLUMNS and 'column' in TOC_COLUMNS['comment'] else None

    toc_list = [toc for toc in toc_list[1:] if toc[process_column] == 'Yes' and toc[level_column] in [0, 1, 2, 3, 4, 5, 6]]

    section_index = 0
    for toc in toc_list:
        data['sections'].append(process_section(context=context, gsheet=gsheet, toc=toc, current_document_index=current_document_index, section_index=section_index, parent=parent, nesting_level=nesting_level))
        section_index = section_index + 1

    return data


def process_section(context, gsheet, toc, current_document_index, section_index, parent, nesting_level):
    # TODO: some columns may have formula, parse those
    # link column (F, toc[link_column] may be a formula), parse it
    if toc[content_type_column] in ['gsheet', 'pdf']:
        link_name, link_target = get_gsheet_link(toc[link_column], nesting_level=nesting_level)
        worksheet_name = link_name

    elif toc[content_type_column] == 'table':
        link_name, link_target = get_worksheet_link(toc[link_column], nesting_level=nesting_level), None
        worksheet_name = link_name

    else:
        link_name, link_target = toc[link_column], None
        worksheet_name = toc[content_type_column]

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
            'label'                 : str(toc[section_column]),
            'heading'               : toc[heading_column],
            'process'               : toc[process_column],
            'level'                 : int(toc[level_column]),
            'content-type'          : toc[content_type_column],
            'link'                  : link_name,
            'link-target'           : link_target,
            'page-break'            : True if toc[break_column] == "page" else False,
            'section-break'         : True if toc[break_column] == "section" else False,
            'page-spec'             : toc[page_spec_column],
            'margin-spec'           : toc[margin_spec_column],

            'landscape'             : True if toc[landscape_column] == "Yes" else False,
			'bookmark'           	: {toc[bookmark_column].strip(): f"{str(toc[section_column])} {toc[heading_column]}".strip()} if bookmark_column is not None else None,
            'autocrop'              : True if toc[autocrop_column] == "Yes" else False,
            'page-bg'               : True if toc[page_bg_column] == "Yes" else False,
            'hide-pageno'           : True if hide_pageno_column is not None and toc[hide_pageno_column] == "Yes" else False,
            'hide-heading'          : True if hide_heading_column is not None and toc[hide_heading_column] == "Yes" else False,
            'different-firstpage'   : True if different_firstpage_column is not None and toc[different_firstpage_column] == "Yes" else False,
            'override-header'       : True if override_header_column is not None and toc[override_header_column] == "Yes" else False,
            'override-footer'       : True if override_footer_column is not None and toc[override_footer_column] == "Yes" else False,
            'background-image'      : toc[background_image_column].strip() if background_image_column is not None else '',
            'responsible'           : toc[responsible_column].strip() if responsible_column is not None else '',
            'reviewer'              : toc[reviewer_column].strip() if reviewer_column is not None else '',
            'status'                : toc[status_column].strip() if status_column is not None else '',
        },
        'header-odd'            : get_worksheet_link(toc[header_odd_column]) if header_odd_column is not None else '',
        'header-even'           : get_worksheet_link(toc[header_even_column]) if header_even_column is not None else '',
        'footer-odd'            : get_worksheet_link(toc[footer_odd_column]) if footer_odd_column is not None else '',
        'footer-even'           : get_worksheet_link(toc[footer_even_column]) if footer_even_column is not None else '',
        'header-first'          : get_worksheet_link(toc[header_first_column]) if header_first_column is not None else '',
        'footer-first'          : get_worksheet_link(toc[footer_first_column]) if footer_first_column is not None else '',
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
