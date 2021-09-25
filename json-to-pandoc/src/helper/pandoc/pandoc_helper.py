#!/usr/bin/env python3

'''
Helper to initialize and manipulate pandoc styles, mostly custom styles not present in the templates
'''

import yaml

from helper.logger import *
from helper.pandoc.pandoc_util import *

HEADER = '''```{=latex}
\\newcommand\Tstrut[1]{\\rule[2.0ex]{0pt}{#1}}                  % "top" strut
\\newcommand\Bstrut[1]{\\rule[-0.9ex]{0pt}{#1}}                 % "bottom" strut
\\newcommand\\titlestyle[1]{{\\vspace{0.2cm}\\begin{center}\Large\\scshape\\textbf #1 \end{center}\\vspace{0.2cm}}}

\\newcommand\\vborder[2]{\\color[rgb]{#1}\\setlength\\arrayrulewidth{#2}\\vline}
\\newcommand\\hborder[2]{\\arrayrulecolor[rgb]{#1}\\setlength\\arrayrulewidth{#2}}

\\setlength{\\tabcolsep}{0.05in}
\\renewcommand{\\arraystretch}{1.2}
```

'''

class PandocHelper(object):

    ''' constructor
    '''
    def __init__(self, style, pandoc_path):
        self._OUTPUT_PANDOC = pandoc_path
        self._STYLE_DATA = style


    ''' initializer - latex YAML header/preamble
    '''
    def init(self):
        self.load_styles()

        self._doc = HEADER

        return self._doc


    ''' save the markdown document string in a file
    '''
    def save(self, doc):
        with open(self._OUTPUT_PANDOC, "w", encoding="utf-8") as f:
            f.write(doc)


    ''' custom styles for sections, etc.
    '''
    def load_styles(self):
        sd = yaml.load(open(self._STYLE_DATA, 'r', encoding='utf-8'), Loader=yaml.FullLoader)

        self._sections = sd['sections']
