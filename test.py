#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import sys
import logging
from package_office.ui import office_ui_main

#Release_ver = True
Release_ver = False

def def_log_module():
    if True == Release_ver:
        loglevel = logging.INFO
        formatter = logging.Formatter('[%(asctime)s]:%(message)s')
    else:
        loglevel = logging.DEBUG
        formatter = logging.Formatter('%(asctime)s-[%(filename)s:%(lineno)s]:%(message)s')
    logger = logging.getLogger("mainModule")
    logger.setLevel(level = loglevel)
    handler = logging.FileHandler("log.txt")
    handler.setLevel(loglevel)
    handler.setFormatter(formatter)

    console = logging.StreamHandler()
    console.setLevel(loglevel)
    console.setFormatter(formatter)

    logger.addHandler(handler)
    logger.addHandler(console)

    logger.info("Applicant Start...")


if __name__ == '__main__':
    office_ui_main()


