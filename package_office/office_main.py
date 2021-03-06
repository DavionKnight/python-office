#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import logging
import shutil
import os.path
import sys
from pathlib import Path
from package_office.txt_oper import do_txt
from package_office.doc_oper import do_doc
from package_office.docx_oper import do_docx

logger = logging.getLogger("mainModule")

txt = ".txt"
doc = ".doc"
docx = ".docx"
input_is_dir = False

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

def is_dir(path):
    return os.path.isdir(path)
def is_file(path):
        return os.path.isfile(path)

def get_file_suffix(path):
    namelist = os.path.splitext(path)
    return namelist[1]
    
def get_file_prefix(path):
    namelist = os.path.splitext(path)
    return namelist[0]

def delWithCmd(path):
    try:
        if os.path.isfile(path):
            cmd = 'del "'+ path + '" /F'
            print(cmd)
            os.system(cmd)
    except Exception as e:
        print(e)

dirsCnt = 0
filesCnt = 0
def deleteDir(dirPath):
    global dirsCnt
    global filesCnt
    for root, dirs, files in os.walk(dirPath, topdown=False):
        for name in files:
            try:
                filesCnt += 1
                filePath = os.path.join(root, name)
                print('file deleted', filesCnt, filePath)
                os.remove(filePath)
            except Exception as e:
                print(e)
                delWithCmd(filePath)
        for name in dirs:
            try:
                os.rmdir(os.path.join(root, name))
                dirsCnt += 1
            except Exception as e:
                print(e)
    print("find file")
    os.rmdir(dirPath)
    print("delete end")

def create_new_dir(path, path_new):
    try:
        if True == is_dir(path_new):
            log_to_ui(True, "info", "?????????????????????????????????!")
            deleteDir(path_new)
        elif True == is_file(path_new):
            delWithCmd(path_new)
    except:
        log_to_ui(True, "info", "??????" + path_new + "?????????????????????????????????")
#        return False
    shutil.copytree(path, path_new)
    log_to_ui(True, "info", "???????????????????????????!??????????????????")
    log_to_ui(True, "", path_new)
    return True

def find_file(arg,dirname,files):
    for file in files:
        file_path=os.path.join(dirname,file)
        if os.path.isfile(file_path):
            print("find file:%s" %file_path)

def save_file(tree, filename='tree.txt'):
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(tree)

tree_str = ''
path_tree_str = ''
customized_tree_str = ''
file_list = []
def generate_tree(path, n=0):
    pathname = Path(path)
    global tree_str
    global path_tree_str
    global customized_tree_str
    global file_list
    next_dir = []
    if pathname.is_file():
        fname = os.path.basename(path)
        #delete file fist char is ~
        if '~' != fname[0]:
            tree_str += '    |' * n + '-' * 4 + pathname.name + '\n'
            path_tree_str += str(pathname.name) + '\n'
            logger.debug(str(pathname.parent.joinpath(pathname.name)))
            file_list.append(str(pathname.parent.joinpath(pathname.name)))
    elif pathname.is_dir():
        tree_str += '    |' * n + '-' * 4 + \
            str(pathname.relative_to(pathname.parent)) + '\\' + '\n'
        logger.debug("??????:" + str(pathname.parent.joinpath(pathname.name)))
        path_tree_str += ('\n??????:' + str(pathname.parent.joinpath(pathname.name)) + '\n')
        customized_tree_str += (str(pathname.parent.joinpath(pathname.name)) + '\n')
        for cp in pathname.iterdir():
            str_cp = str(cp)
            if cp.is_file():
                generate_tree(str_cp, n + 1)
            elif cp.is_dir():
                next_dir.append(str_cp)
        next_dir.sort()
        for cp1 in next_dir:
            generate_tree(cp1, n + 1)

def dir_tree_generate(path):
    log_to_ui(True, "info", "??????????????????...")
    global tree_str
    global path_tree_str
    global customized_tree_str
    global file_list
    tree_str = ''
    path_tree_str = ''
    customized_tree_str = ''
    file_list = []
    generate_tree(path)
    logger.info('\n????????????\n' + tree_str)
    logger.info('\n????????????\n' + path_tree_str)
    from package_office.ui import ui_dir_tree_show
    ui_dir_tree_show(tree_str)
    save_file(tree_str)
    save_file(path_tree_str, 'path_tree_str.txt')
    save_file(customized_tree_str, 'customized_tree_str.txt')
    log_to_ui(True, "info", "?????????????????????!")
#    for root,dirs,files in os.walk(path, topdown=False):
#        dirs.sort()
#        for dir in dirs:
#            logger.debug("##??????:" + os.path.join(root,dir).encode('utf-8').decode('gbk'))
#        for file in files:
#            logger.debug(os.path.join(root,file).encode('utf-8').decode('gbk'))

def rename_file(path, replace_dict):
    log_to_ui(True, "info", "??????????????????...")
    for parent,dirnames,filenames in os.walk(path):#???????????????????????????1.????????? 2.??????????????????????????????????????? 3.??????????????????
        for filename in filenames:#?????????
            for key, value in replace_dict.items():
                if key in filename:
                    nfn = filename.replace(key, value)
                    os.rename(os.path.join(parent,filename),os.path.join(parent,nfn))#?????????
                    log_to_ui(True, "info", "??????" + filename + "????????????" + nfn)
    log_to_ui(True, "info", "????????????????????????")

def dir_traversal_opt(path, replace_dict):
    global file_list
    for fname in file_list:
        file_opt(fname, replace_dict)

def dir_opt(path, replace_dict):
    log_to_ui(True, "info", "???????????????????????????")
    log_to_ui(True, "", path)

    global input_is_dir
    input_is_dir = True

    path_new = os.path.dirname(path) + "/" + os.path.basename(path) + "_new"
    if False == create_new_dir(path, path_new):
        return False

    rename_file(path_new, replace_dict)

    dir_tree_generate(path_new, )

    dir_traversal_opt(path_new, replace_dict)

def gen_newpath(path):
    global input_is_dir
    if True == input_is_dir:
        name = path
    else:
        name = get_file_prefix(path)
        name = name + "_new"
        name = name + get_file_suffix(path)
    return name

def file_opt(path, replace_dict):
    log_to_ui(True, "", "")
    log_to_ui(True, "info", "?????????????????????" + path)

    fname = os.path.basename(path)
    if '~' == fname[0]:
        logger.error("??????" + fsuffix + "???????????????????????????!")
        return False

    fsuffix = get_file_suffix(path)
    new_name = gen_newpath(path)
    logger.info("File is:" + new_name)
    if txt == fsuffix:
        do_txt(path, new_name, replace_dict)
    elif doc == fsuffix:
        do_doc(path, new_name, replace_dict)
    elif docx == fsuffix:
        do_docx(path, new_name, replace_dict)
    else:
        log_to_ui(True, "error", "??????" + path + "???????????????????????????")
        return
    log_to_ui(True, "info", "???????????????")

def office_main(input_path, replace_dict):
    if True == is_dir(input_path):
        if False == dir_opt(input_path, replace_dict):
            return False
    elif True == is_file(input_path):
        file_opt(input_path, replace_dict)
    else:
        log_to_ui(True, "error", "file name error!")
    log_to_ui(True, "", "")
    log_to_ui(True, "info", "????????????????????????!")

