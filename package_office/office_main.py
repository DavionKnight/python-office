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

def create_new_dir(path, path_new):
    try:
        if True == is_dir(path_new):
            shutil.rmtree(path_new)
        elif True == is_file(path_new):
            os.remove(path_new)
        else:
            logger.info("create path_new:" + path_new)
    except:
        logger.error(" repr(e):", repr(e))
    shutil.copytree(path, path_new)

def dir_traversal_opt(path, replace_dict):
    for fname in os.listdir(path):
        logger.debug("list:" + fname)
        fname_new = os.path.join(path, fname)
        file_opt(fname_new, replace_dict)

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
def generate_tree(path, n=0):
    pathname = Path(path)
    global tree_str
    global path_tree_str
    global customized_tree_str
    next_dir = []
    if pathname.is_file():
        tree_str += '    |' * n + '-' * 4 + pathname.name + '\n'
        path_tree_str += str(pathname.name) + '\n'
        logger.info(str(pathname.parent.joinpath(pathname.name)))
    elif pathname.is_dir():
        tree_str += '    |' * n + '-' * 4 + \
            str(pathname.relative_to(pathname.parent)) + '\\' + '\n'
        logger.info("目录:" + str(pathname.parent.joinpath(pathname.name)))
        path_tree_str += ('\n目录:' + str(pathname.parent.joinpath(pathname.name)) + '\n')
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

def dir_show(path):
    generate_tree(path)
    logger.debug(tree_str)
    save_file(tree_str)
    save_file(path_tree_str, 'path_tree_str.txt')
    save_file(customized_tree_str, 'customized_tree_str.txt')
    for root,dirs,files in os.walk(path, topdown=False):
        dirs.sort()
        for dir in dirs:
            logger.debug("##目录:" + os.path.join(root,dir).encode('utf-8').decode('gbk'))
        for file in files:
            logger.debug(os.path.join(root,file).encode('utf-8').decode('gbk'))

def dir_opt(path, replace_dict):
    logger.info("DIR:" + path)
    dir_show(path)
    path_new = os.path.dirname(path) + os.path.basename(path) + "_new"
    create_new_dir(path, path_new)
    input_is_dir = True
    dir_traversal_opt(path_new, replace_dict)
    logger.info(path_new + " has interatored done!")

def gen_newpath(path):
    if True == input_is_dir:
        name = path
    else:
        name = get_file_prefix(path)
        name = name + "_new"
        name = name + get_file_suffix(path)
    logger.info("new file is:" + name)
    return name

def file_opt(path, replace_dict):
    logger.debug(path + " is file")
    fsuffix = get_file_suffix(path)
    new_name = gen_newpath(path)
    if txt == fsuffix:
        do_txt(path, new_name, replace_dict)
    elif doc == fsuffix:
        do_doc(path, new_name, replace_dict)
    elif docx == fsuffix:
        do_docx(path, new_name, replace_dict)
    else:
        logger.error("ERROR:file suffix can not recongnized! suffix get is:" + fsuffix)

def office_main(input_path, replace_dict):
    if True == is_dir(input_path):
#        input_is_dir = True
        dir_opt(input_path, replace_dict)
    elif True == is_file(input_path):
#        input_is_dir = False
        file_opt(input_path, replace_dict)
    else:
        logger.error ("file name error!")

