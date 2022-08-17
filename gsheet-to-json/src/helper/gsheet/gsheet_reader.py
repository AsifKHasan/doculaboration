#!/usr/bin/env python3
'''
'''
from collections import defaultdict

import re
import importlib

# import pandas as pd
import pygsheets

import urllib.request

from helper.logger import *
from helper.gsheet.gsheet_util import *


def process_gsheet(context, sheet, parent=None):
    data = {}

    # worksheet-cache is nested dictionary of sheet->worksheet as two different sheets may have worksheets of same name
    # so keying by only worksheet name is not feasible
    if 'worksheet-cache' not in context:
        context['worksheet-cache'] = {}

    if sheet.title not in context['worksheet-cache']:
        context['worksheet-cache'][sheet.title] = {}

    ws_title = context['index-worksheet']
    ws = sheet.worksheet('title', ws_title)

    toc_list = ws.get_values(start='A3', end=f"X{ws.rows}", returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=False, value_render='FORMULA')
    toc_list = [toc for toc in toc_list if toc[2] == 'Yes' and toc[3] in [0, 1, 2, 3, 4, 5, 6]]

    data['sections'] = [process_section(context, sheet, toc, parent) for toc in toc_list]

    return data


def process_section(context, sheet, toc, parent=None):
    # TODO: some columns may have formula, parse those
    # link column (F, toc[5] may be a formula), parse it
    if toc[4] in ['gsheet', 'docx', 'pdf']:
        link_name, link_target = get_gsheet_link(toc[5])

    elif toc[4] == 'table':
        link_name, link_target = get_worksheet_link(toc[5]), None

    else:
        link_name, link_target = toc[5], None


    # transform to a dict
    d = {
        'section'               : str(toc[0]),
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
        'header-first'          : get_worksheet_link(toc[13]),
        'header-odd'            : get_worksheet_link(toc[14]),
        'header-even'           : get_worksheet_link(toc[15]),
        'footer-first'          : get_worksheet_link(toc[16]),
        'footer-odd'            : get_worksheet_link(toc[17]),
        'footer-even'           : get_worksheet_link(toc[18]),
        'override-header'       : True if toc[19] == "Yes" else False,
        'override-footer'       : True if toc[20] == "Yes" else False,
        'responsible'           : toc[21],
        'reviewer'              : toc[22],
        'status'                : toc[23]
        }

    # the gsheet is a child gsheet, called from a parent gsheet, so header processing depends on override flags
    if parent:
        if parent['override-header']:
            # debug(f".. this is a child gsheet : header OVERRIDDEN")
            d['different-firstpage'] = parent['different-firstpage']
            d['header-first'] = parent['header-first']
            d['header-odd'] = parent['header-odd']
            d['header-even'] = parent['header-even']
        else:
            # debug(f".. child gsheet's header is NOT overridden")
            module = importlib.import_module('processor.table_processor')
            if d['different-firstpage']:
                if d['header-first'] != '' and d['header-first'] is not None:
                    d['header-first'] = module.process(sheet, {'link': d['header-first']}, context)
                else:
                    d['header-first'] = None
            else:
                d['header-first'] = None

            if d['header-odd'] != '' and d['header-odd'] is not None:
                d['header-odd'] = module.process(sheet, {'link': d['header-odd']}, context)
            else:
                d['header-odd'] = None

            if d['header-even'] != '' and d['header-even'] is not None:
                d['header-even'] = module.process(sheet, {'link': d['header-even']}, context)
            else:
                d['header-even'] = d['header-odd']

        if parent['override-footer']:
            # debug(f".. this is a child gsheet : footer OVERRIDDEN")
            d['different-firstpage'] = parent['different-firstpage']
            d['footer-first'] = parent['footer-first']
            d['footer-odd'] = parent['footer-odd']
            d['footer-even'] = parent['footer-even']
        else:
            # debug(f".. child gsheet's footer is NOT overridden")
            module = importlib.import_module('processor.table_processor')
            if d['different-firstpage']:
                if d['footer-first'] != '' and d['footer-first'] is not None:
                    d['footer-first'] = module.process(sheet, {'link': d['footer-first']}, context)
                else:
                    d['footer-first'] = None
            else:
                d['footer-first'] = None

            if d['footer-odd'] != '' and d['footer-odd'] is not None:
                d['footer-odd'] = module.process(sheet, {'link': d['footer-odd']}, context)
            else:
                d['footer-odd'] = None

            if d['footer-even'] != '' and d['footer-even'] is not None:
                d['footer-even'] = module.process(sheet, {'link': d['footer-even']}, context)
            else:
                d['footer-even'] = d['footer-odd']
    else:
        module = importlib.import_module('processor.table_processor')

        # process header, it may be text or link
        if d['different-firstpage']:
            if d['header-first'] != '' and d['header-first'] is not None:
                d['header-first'] = module.process(sheet, {'link': d['header-first']}, context)
            else:
                d['header-first'] = None

            if d['footer-first'] != '' and d['footer-first'] is not None:
                d['footer-first'] = module.process(sheet, {'link': d['footer-first']}, context)
            else:
                d['footer-first'] = None
        else:
            d['header-first'] = None
            d['footer-first'] = None

        if d['header-odd'] != '' and d['header-odd'] is not None:
            d['header-odd'] = module.process(sheet, {'link': d['header-odd']}, context)
        else:
            d['header-odd'] = None

        if d['header-even'] != '' and d['header-even'] is not None:
            d['header-even'] = module.process(sheet, {'link': d['header-even']}, context)
        else:
            d['header-even'] = d['header-odd']

        if d['footer-odd'] != '' and d['footer-odd'] is not None:
            d['footer-odd'] = module.process(sheet, {'link': d['footer-odd']}, context)
        else:
            d['footer-odd'] = None

        if d['footer-even'] != '' and d['footer-even'] is not None:
            d['footer-even'] = module.process(sheet, {'link': d['footer-even']}, context)
        else:
            d['footer-even'] = d['footer-odd']

    # import and use the specific processor
    if d['link'] == '' or d['link'] is None:
        d['contents'] = None
    else:
        module = importlib.import_module(f"processor.{d['content-type']}_processor")
        d['contents'] = module.process(sheet, d, context)

    return d
