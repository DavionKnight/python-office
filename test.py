#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import logging
from package_office.office_main import office_main

Release_ver = True
#Release_ver = False

text_path = "./aa/bb.txt"
text_ori = "一城龙域"
text_dest = "二城龙域"
text_ori1 = "认证"
text_dest1 = "我靠"
replace_dict1 = {
        text_ori:text_dest,
        text_ori1:text_dest1
}

docx_path = "./aa/docx.docx"
docx_ori = "参数1"
docx_dest = "XXXXXX1"
docx_ori1 = "参数2"
docx_dest1 = "QQQQQ2"
replace_dict2 = {
        docx_ori:docx_dest,
        docx_ori1:docx_dest1
}

dir_path = "./test/"

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
    def_log_module()
    input_path = dir_path
    office_main(input_path, replace_dict2)
#    input_path = docx_path
#    office_main(input_path, replace_dict2)
#    input_path = text_path
#    office_main(input_path, replace_dict1)


