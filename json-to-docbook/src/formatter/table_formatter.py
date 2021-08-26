#!/usr/bin/env python3

import textwrap
import importlib

from helper.logger import *
from helper.docbook.docbook_writer import *
from helper.docbook.docbook_util import *

def generate(data, doc, section_specs, context):
    error('formatter [table] not supported')
