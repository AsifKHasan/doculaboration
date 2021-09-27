#!/usr/bin/env python3

'''
various utilities for generating latex code
'''

import re

CONV = {
    '&': r'\&',
    '%': r'\%',
    '$': r'\$',
    '#': r'\#',
    '_': r'\_',
    '{': r'\{',
    '}': r'\}',
    '~': r'\textasciitilde{}',
    '^': r'\^{}',
    '\\': r'\textbackslash{}',
    '<': r'\textless{}',
    '>': r'\textgreater{}',
}

GSHEET_OXML_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dashed',
    'SOLID': 'single',
    'SOLID_MEDIUM': 'thick',
    'SOLID_THICK': 'triple',
    'DOUBLE': 'double',
    'NONE': 'none'
}

VALIGN = {'TOP': 'p', 'MIDDLE': 'm', 'BOTTOM': 'b'}
HALIGN = {'LEFT': '\\raggedright', 'CENTER': '\centering', 'RIGHT': '\\raggedleft', 'JUSTIFY': ''}


'''
'''
def begin_latex():
    return "```{=latex}"


def end_latex():
    return "```\n\n"


''' :param text: a plain text message
    :return: the message escaped to appear correctly in LaTeX
'''
def tex_escape(text):
    regex = re.compile('|'.join(re.escape(str(key)) for key in sorted(CONV.keys(), key = lambda item: - len(item))))
    return regex.sub(lambda match: CONV[match.group()], text)
