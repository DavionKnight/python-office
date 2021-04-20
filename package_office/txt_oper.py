#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import logging

logger = logging.getLogger("mainModule")

def check_all_match_count(fstr, replace_dict):
    count = 0

    new_str = fstr.replace(' ', '').replace('\r', '').replace('\n', '')
    for key, value in replace_dict.items():
        if key in fstr:
            count += 1
    return count

def do_txt(old_path, new_path, replace_dict):
    logger.debug("do txt")
    count_logical = 0
    count_real = 0
    try:
        file_ori = open(old_path,'r') #打开所有文件
        fstr = file_ori.read()
        file_ori.close()
    except:
        logger.error ("open " + old_path + " failed")
        return

    count_logical = check_all_match_count(fstr, replace_dict)
    for key, value in replace_dict.items():
        if key in fstr:
            fstr = fstr.replace(key, value)
            count_real += 1
            logger.info("replace " + key + " to " + value)
    try:
        if 0 != count_real:
            file_dest = open(new_path,'w')
            file_dest.write(fstr) #再写入
            file_dest.close() #关闭文件
    except:
        logger.error("write" + new_path + " failed")
        return
    if count_logical == count_real:
        logger.info("--->File " + old_path + " has checked same words count:%d" %count_logical)
    else:
        logger.info("--->Warnning! " + old_path + " has checked same words:%d" % count_logical + " but replaced count:%d" % count_real)

    if 0 == count_real:
        logger.info("--->Nothing to do, file " + old_path)
    else:
        logger.info("--->replace done, new file " + new_path)
