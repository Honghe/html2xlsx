# -*- coding:utf-8 -*-
import logging
import logging.config
import os

import yaml

_init = False


def get_logger():
    global _init
    if not _init:
        with open(os.path.join(os.path.normpath(os.path.dirname(__file__)), 'logconfig.yaml'), 'r') as f:
            config = yaml.safe_load(f.read())
            logging.config.dictConfig(config)
        _init = True
    return logging.getLogger()