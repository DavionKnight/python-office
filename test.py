#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import sys
import logging
from package_office.office_main import office_main

#Release_ver = True
Release_ver = False

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
docx_ori2 = "初始版本"
docx_dest2 = "最后版本"
replace_dict2 = {
        docx_ori:docx_dest,
        docx_ori1:docx_dest1,
        docx_ori2:docx_dest2
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
    print('参数个数为:', len(sys.argv), '个参数。')
    print('参数列表:', str(sys.argv))
    print('目标目录:', (sys.argv[1]))
    print('原文本:', (sys.argv[2]))
    print('替换为:', (sys.argv[3]))
#    input_path = dir_path
    replace_dict = {sys.argv[2]:sys.argv[3]}
    #print('Item:', sys.argv[2], 'value:' , (sys.argv[3]))
    office_main(sys.argv[1], replace_dict)
#    input_path = docx_path
#    office_main(input_path, replace_dict2)
#    input_path = text_path
#    office_main(input_path, replace_dict1)


