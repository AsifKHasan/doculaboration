#!/usr/bin/env python3

import time
from datetime import datetime
from termcolor import colored
import colorama

# log level (0-TRACE, 1-DEBUG, 2-INFO, 3-WARN, 4-ERROR) below which logs will not be printed
LOG_LEVEL = None

colorama.init()

log_color = {
    '[ERROR]': {'color': 'red',         'highlight': None, 'attrs': ['bold']},
    '[ WARN]': {'color': 'yellow',      'highlight': None, 'attrs': ['bold']},
    '[ INFO]': {'color': 'white',       'highlight': None, 'attrs': None},
    '[DEBUG]': {'color': 'green',       'highlight': None, 'attrs': ['dark']},
    '[TRACE]': {'color': 'light_grey',  'highlight': None, 'attrs': ['dark']}
}

def trace(msg, console=True, nesting_level=0):
    if LOG_LEVEL < 1:
        log('[TRACE]', msg, console, nesting_level)

def debug(msg, console=True, nesting_level=0):
    if LOG_LEVEL < 2:
        log('[DEBUG]', msg, console, nesting_level)

def info(msg, console=True, nesting_level=0):
    if LOG_LEVEL < 3:
        log('[ INFO]', msg, console, nesting_level)

def warn(msg, console=True, nesting_level=0):
    if LOG_LEVEL < 4:
        log('[ WARN]', msg, console, nesting_level)

def error(msg, console=True, nesting_level=0):
    if LOG_LEVEL >= 4:
        log('[ERROR]', msg, console, nesting_level)

def log(level, msg, console=True, nesting_level=0):
    now = time.time()
    nesting_leader = ".." * nesting_level
    if nesting_leader != '':
        nesting_leader = nesting_leader + ' '

    data = {'type': level, 'time': datetime.now().isoformat(), 'msg': f"{nesting_leader}{msg}"}

    if console:
        text = f"{data['time']} {data['type']:<6} {data['msg']}"
        print(colored(text, log_color[data['type']]['color'], log_color[data['type']]['highlight'], attrs=log_color[data['type']]['attrs']))
