#!/usr/bin/python

import logging
from logging.handlers import RotatingFileHandler

fh = RotatingFileHandler("log/app.log", maxBytes=10 * 1024 * 1024, backupCount=100)
fh.setLevel(logging.DEBUG)
#log write in console
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
#log formatter
formatter = logging.Formatter(
    '%(asctime)s %(levelname)8s [%(filename)s%(lineno)06s] %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)

logger = logging.root
logger.setLevel(logging.INFO)
logger.addHandler(fh)
logger.addHandler(ch)
