# -*- coding: utf-8 -*-

import logging

_log_level = logging.INFO
#_log_level = logging.DEBUG

def _getConsoleHandler():
    h = logging.StreamHandler()
    h.setFormatter(logging.Formatter('%(asctime)s %(name)s:%(lineno)s %(funcName)s [%(levelname)s]: %(message)s'))
    return h

def getLogger(name):
    l = logging.getLogger(name)
    l.setLevel(_log_level)
    l.addHandler(_getConsoleHandler())
    return l
#[EOF]