#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import logging
import pythoncom
import shutil
from win32com import client as wc
from package_office.docx_oper import do_docx

logger = logging.getLogger("mainModule")

def log_to_ui(to_ui, level, msg):
    msg_level = ""
    if "debug" == level:
        msg_level = "[DEBUG]"
        logger.debug(msg_level + msg)
    elif "error" == level:
        msg_level = "[ERROR]"
        logger.error(msg_level + msg)
    elif "info" == level:
        msg_level = "[INFO]"
        logger.info(msg_level + msg)
    else:
        logger.debug("[DEBUG]" + msg)
    if True == to_ui:
        from package_office.ui import ui_log_show
        ui_log_show(msg_level + msg)

def do_doc(path, new_path, replace_dict):
    print ("do doc")
    pythoncom.CoInitialize()
    try:
        w = wc.Dispatch('kwps.Application')
    except:
        w = wc.Dispatch('wps.Application')
    finally:
        doc = w.Documents.Open(path)
        new_filepath = os.path.splitext(path)[0]+".docx"
        doc.SaveAs(new_filepath, 12, False, "", True, "", False, False, False,False)  # 转化后路径下的文件
        doc.Close()
        w.Quit()
#    shutil.rmtree(path)
        print ("new_filepath" + new_filepath)
        from package_office.office_main import delWithCmd
        delWithCmd(path)
        log_to_ui(True, "info", "文件" + new_path + "转换为" + new_filepath)
        do_docx(new_filepath, new_filepath, replace_dict)
