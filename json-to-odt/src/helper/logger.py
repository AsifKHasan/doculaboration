import time
from datetime import datetime
from termcolor import colored
import colorama

colorama.init()

log_color = {
    '[ERROR]': {'color': 'red',    'highlight': None, 'attrs': ['bold']},
    '[ WARN]':  {'color': 'yellow', 'highlight': None, 'attrs': ['bold']},
    '[ INFO]':  {'color': 'white',  'highlight': None, 'attrs': None},
    '[DEBUG]': {'color': 'green',  'highlight': None, 'attrs': None}
}

def debug(msg, console=True):
    log('[DEBUG]', msg, console)

def info(msg, console=True):
    log('[ INFO]', msg, console)

def warn(msg, console=True):
    log('[ WARN]', msg, console)

def error(msg, console=True):
    log('[ERROR]', msg, console)

def log(level, msg, console=True):
    now = time.time()
    data = {'type': level, 'time': datetime.now().isoformat(), 'msg': msg}

    if console:
        text = '{0} {1:<6} {2}'.format(data['time'], data['type'], data['msg'])
        # print(text)
        print(colored(text, log_color[data['type']]['color'], log_color[data['type']]['highlight'], attrs=log_color[data['type']]['attrs']))
