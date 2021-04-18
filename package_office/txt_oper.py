#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import logging

logger = logging.getLogger("mainModule")
def do_txt(old_path, new_path, replace_dict):
    logger.debug("do txt")
    try:
        file_ori = open(old_path,'r') #打开所有文件
        str = file_ori.read()
        file_ori.close()
    except:
        logger.error ("open " + old_path + " failed")
        return

    for key, value in replace_dict.items():
        if key in str:
            str = str.replace(key, value)
            logger.info("replace " + key + " to " + value)
            
    try:
        file_dest = open(new_path,'w')
        file_dest.write(str) #再写入
    except:
        logger.error("write" + new_path + " failed")
        return
    file_dest.close() #关闭文件
    logger.info("--->replace done, new file " + new_path)
