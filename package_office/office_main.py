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

def create_new_dir(path, path_new):
    try:
        if True == is_dir(path_new):
            log_to_ui(True, "info", "新目录存在，删除新目录!")
            shutil.rmtree(path_new)
        elif True == is_file(path_new):
            os.remove(path_new)
    except:
        logger.error(" repr(e):", repr(e))
    shutil.copytree(path, path_new)
    log_to_ui(True, "info", "新目录重新生成完成!，新目录为：")
    log_to_ui(True, "", path_new)

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
        logger.debug("目录:" + str(pathname.parent.joinpath(pathname.name)))
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

def dir_tree_generate(path):
    log_to_ui(True, "info", "目录树生成中...")
    global tree_str
    global path_tree_str
    global customized_tree_str
    global file_list
    tree_str = ''
    path_tree_str = ''
    customized_tree_str = ''
    file_list = []
    generate_tree(path)
    logger.info('\n目录树：\n' + tree_str)
    logger.info('\n目录树：\n' + path_tree_str)
    from package_office.ui import ui_dir_tree_show
    ui_dir_tree_show(tree_str)
    save_file(tree_str)
    save_file(path_tree_str, 'path_tree_str.txt')
    save_file(customized_tree_str, 'customized_tree_str.txt')
    log_to_ui(True, "info", "目录树生成完成!")
#    for root,dirs,files in os.walk(path, topdown=False):
#        dirs.sort()
#        for dir in dirs:
#            logger.debug("##目录:" + os.path.join(root,dir).encode('utf-8').decode('gbk'))
#        for file in files:
#            logger.debug(os.path.join(root,file).encode('utf-8').decode('gbk'))

def rename_file(path, replace_dict):
    log_to_ui(True, "info", "文件重命名中...")
    for parent,dirnames,filenames in os.walk(path):#三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:#文件名
            for key, value in replace_dict.items():
                if key in filename:
                    nfn = filename.replace(key, value)
                    os.rename(os.path.join(parent,filename),os.path.join(parent,nfn))#重命名
                    log_to_ui(True, "info", "文件" + filename + "重命名为" + nfn)
    log_to_ui(True, "info", "文件重命名完成。")

def dir_traversal_opt(path, replace_dict):
    global file_list
    for fname in file_list:
        file_opt(fname, replace_dict)

def dir_opt(path, replace_dict):
    log_to_ui(True, "info", "您要替换的目录为：")
    log_to_ui(True, "", path)

    global input_is_dir
    input_is_dir = True

    path_new = os.path.dirname(path) + "/" + os.path.basename(path) + "_new"
    create_new_dir(path, path_new)

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
    log_to_ui(True, "info", "开始处理文件：" + path)

    fname = os.path.basename(path)
    if '~' == fname[0]:
        logger.error("文件" + fsuffix + "不是文本文件，跳过!")
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
        log_to_ui(True, "error", "文件" + path + "后缀不识别，跳过！")
        return
    log_to_ui(True, "info", "文件" + path + "处理完成！")

def office_main(input_path, replace_dict):
    if True == is_dir(input_path):
        dir_opt(input_path, replace_dict)
    elif True == is_file(input_path):
        file_opt(input_path, replace_dict)
    else:
        log_to_ui(True, "error", "file name error!")
    log_to_ui(True, "", "")
    log_to_ui(True, "info", "文件操作全部完成!")

