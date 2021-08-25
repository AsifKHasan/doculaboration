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

def process_sheet(context, sheet, parent=None):
    data = {}

    # worksheet-cache is nested dictionary of sheet->worksheet as two different sheets may have worksheets of same name, so keying by only worksheet name is not feasible
    if 'worksheet-cache' not in context:
        context['worksheet-cache'] = {}

    if sheet.title not in context['worksheet-cache']:
        context['worksheet-cache'][sheet.title] = {}

    ws_title = context['index-worksheet']
    ws = sheet.worksheet('title', ws_title)
    toclist = ws.get_values(start='A3', end='U{}'.format(ws.rows), returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=False)
    toclist = [toc for toc in toclist if toc[2] == 'Yes' and toc[3] in ['0', '1', '2', '3', '4', '5', '6']]
    data['sections'] = [process_section(sheet, s, context, parent) for s in toclist]
    return data

def process_section(sheet, s, context, parent=None):
    # transform to a dict
    d = {
        'section': str(s[0]), 'heading': s[1], 'process': s[2], 'level': int(s[3]), 'content-type': s[4], 'link': s[5], 'section-break': s[6],
        'page-number': False if s[7] == "No" else True,
        'no-heading': True if s[8] == "Yes" else False,
        'different-first-page-header-footer': True if s[9] == "Yes" else False,
        'header-first': s[10], 'header-odd': s[11], 'header-even': s[12],
        'footer-first': s[13], 'footer-odd': s[14], 'footer-even': s[15],
        'override-child-header': True if s[16] == "Yes" else False,
        'override-child-footer': True if s[17] == "Yes" else False,
        'responsible': s[18], 'reviewer': s[19], 'status': s[20]
        }

    # the gsheet is a child gsheet, called from a parent gsheet, so header processing depends on override flags
    if parent:
        warn('This is a child gsheet')

        if parent['override-child-header']:
            warn('child gsheet''s header is OVERRIDDEN')
            d['different-first-page-header-footer'] = parent['different-first-page-header-footer']
            d['header-first'] = parent['header-first']
            d['header-odd'] = parent['header-odd']
            d['header-even'] = parent['header-even']
        else:
            warn('child gsheet''s header is NOT overridden')
            module = importlib.import_module('processor.table_processor')
            if d['different-first-page-header-footer']:
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

        if parent['override-child-footer']:
            warn('child gsheet''s footer is OVERRIDDEN')
            d['different-first-page-header-footer'] = parent['different-first-page-header-footer']
            d['footer-first'] = parent['footer-first']
            d['footer-odd'] = parent['footer-odd']
            d['footer-even'] = parent['footer-even']
        else:
            warn('child gsheet''s footer is NOT overridden')
            module = importlib.import_module('processor.table_processor')
            if d['different-first-page-header-footer']:
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
        if d['different-first-page-header-footer']:
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
        module = importlib.import_module('processor.{0}_processor'.format(d['content-type']))
        d['contents'] = module.process(sheet, d, context)

    return d
