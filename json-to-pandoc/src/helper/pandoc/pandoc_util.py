#!/usr/bin/env python3

'''
various utilities for generating a pandoc markdown document
'''

import sys
from helper.logger import *

''' :param path: a path string
    :return: the path that the OS accepts
'''
def os_specific_path(path):
    if sys.platform == 'win32':
        return path.replace('\\', '/')
    else:
        return path
