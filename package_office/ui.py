#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import logging
from package_office.office_main import office_main

from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter import StringVar
import time

import matplotlib.pyplot as plt

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

LOG_LINE_NUM = 0

class MY_GUI():

    def __init__(self):
        self.tk = Tk()              #实例化出一个父窗口
        self.tk.configure(bg='Lavender') 
        self.path = StringVar()
        self.dir_tree = StringVar()
        self.o1 = StringVar()
        self.o2 = StringVar()
        self.o3 = StringVar()
        self.d1 = StringVar()
        self.d2 = StringVar()
        self.d3 = StringVar()
        self.tk.title("文本替换工具_v1.0 by meteor(V:kevin_fzs)")           #窗口名
        #self.tk.geometry('320x160+10+10')                         #290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.tk.geometry('1200x800+10+10')

    def selectPath(self):
        path_ = askdirectory()
        self.path.set(path_)
#        print('参数个数为:', len(sys.argv), '个参数。')
#        print('参数列表:', str(sys.argv))
#        print('目标目录:', (sys.argv[1]))
#        print('原文本:', (sys.argv[2]))
#        print('替换为:', (sys.argv[3]))
#        replace_dict = {sys.argv[2]:sys.argv[3]}
        #print('Item:', sys.argv[2], 'value:' , (sys.argv[3]))
#        office_main(sys.argv[1], replace_dict)

    def office_test(self):
        self.dir_tree_text.delete('1.0', END)

    def log_show(self, msg):
        self.log_text.insert(END, msg)
        self.log_text.insert(INSERT, '\n')

    def office_execute(self):
        self.log_text.delete('1.0', END)
        print('o1:', self.o1.get())
        print('d1:', self.d1.get())
        print('o2:', self.o2.get())
        print('d2:', self.d2.get())
        print('o3:', self.o3.get())
        print('d3:', self.d3.get())
        dict = {}
        dict.clear()
        if not (len(self.o1.get()) == 0 or len(self.d1.get()) == 0):
            dict[self.o1.get()] = self.d1.get()
            print('dict[o1]:', dict[self.o1.get()])
        if not (len(self.o2.get()) == 0 or len(self.d2.get()) == 0):
            dict[self.o2.get()] = self.d2.get()
            print('dict[o2]:', dict[self.o2.get()])
        if not (len(self.o3.get()) == 0 or len(self.d3.get()) == 0):
            dict[self.o3.get()] = self.d3.get()
            print('dict[o3]:', dict[self.o3.get()])

        print('dict len:', len(dict))
        if len(self.path.get()) == 0:
            print('ERROR:Please select path')
            self.log_show("请选择要处理的目录！")
            return
        if len(dict) == 0:
            self.log_show("请输入要替换的内容，至少输入一行！")
            print('ERROR:Please input replace words')
            return

        office_main(self.path.get(), dict)

    #设置窗口
    def set_init_window(self):
        #self.tk["bg"] = "pink"                                    #窗口背景色，其他背景色见：blog.csdn.net/chl0000/article/details/7657887
        #self.tk.attributes("-alpha",0.9)                          #虚化，值越小虚化程度越高
        #标签
        Label(self.tk, text="请选择处理目录:").grid(row=0, column=0,sticky=W)
        Entry(self.tk, textvariable = self.path).grid(row = 1, column = 0)
        Button(self.tk, text = "路径选择", command = self.selectPath).grid(row = 1, column = 1)
        Button(self.tk, text = "执行", command = self.office_execute).grid(row = 1, column = 2)
        Button(self.tk, text = "test", command = self.office_test).grid(row = 2, column = 2)

        Label(self.tk, text="文件内容和文件名替换):").grid(row=2, column=0,sticky=W)
        Entry(self.tk, textvariable=self.o1).grid(row = 4, column = 0)
        Label(self.tk, text="替换为").grid(row = 4, column = 1)
        Entry(self.tk, textvariable=self.d1).grid(row = 4, column = 2)

        Entry(self.tk, textvariable=self.o2).grid(row = 5, column = 0)
        Label(self.tk, text="替换为").grid(row = 5, column = 1)
        Entry(self.tk, textvariable=self.d2).grid(row = 5, column = 2)

        entry_usr_name = Entry(self.tk, textvariable=self.o3).grid(row = 6, column = 0)
        Label(self.tk, text="替换为").grid(row = 6, column = 1)
        Entry(self.tk, textvariable=self.d3).grid(row = 6, column = 2)

        Label(self.tk, text="目录树").grid(row = 7, column = 0,sticky=W)
        self.dir_tree_text = Text(self.tk, width=60, height=40)  #原始数据录入框
        self.dir_tree_text.grid(row=8, column=0, rowspan=5, columnspan=3,sticky=W)

        Label(self.tk, text="日志:").grid(row=0, column=4,sticky=W)
        self.log_text = Text(self.tk, width=100, height=50)  #原始数据录入框
        self.log_text.grid(row=1, column=4, rowspan=60, columnspan=60, sticky=N+W)
#        plt.grid(linestyle=":", color="r")
#        w = Canvas(self.tk, width=200, height=200, background="white" )
#        w.pack()
#        w.create_line(100, 0,100, 200,fill='red', dash=(4, 4))
#        self.result_data_label = Label(self.tk, text="目录树")
#        self.result_data_label.grid(row=2, column=12)
#        self.log_label = Label(self.tk, text="日志")
#        self.log_label.grid(row=12, column=0)
#        #文本框
#        self.init_data_Text = Text(self.tk, width=67, height=35)  #原始数据录入框
#        self.init_data_Text.grid(row=2, column=0, rowspan=10, columnspan=10)
#        self.result_data_Text = Text(self.tk, width=70, height=49)  #处理结果展示
#        self.result_data_Text.grid(row=1, column=12, rowspan=15, columnspan=10)
#        self.log_data_Text = Text(self.tk, width=66, height=9)  # 日志框
#        self.log_data_Text.grid(row=13, column=0, columnspan=10)
#        #按钮
#        self.str_trans_to_md5_button = Button(self.tk, text="字符串转MD5", bg="lightblue", width=10,command=self.str_trans_to_md5)  # 调用内部方法  加()为直接调用
#        self.str_trans_to_md5_button.grid(row=1, column=11)


    #功能函数
    def str_trans_to_md5(self):
        src = self.init_data_Text.get(1.0,END).strip().replace("\n","").encode()
        #print("src =",src)
        if src:
            try:
                myMd5 = hashlib.md5()
                myMd5.update(src)
                myMd5_Digest = myMd5.hexdigest()
                #print(myMd5_Digest)
                #输出到界面
                self.result_data_Text.delete(1.0,END)
                self.result_data_Text.insert(1.0,myMd5_Digest)
                self.write_log_to_Text("INFO:str_trans_to_md5 success")
            except:
                self.result_data_Text.delete(1.0,END)
                self.result_data_Text.insert(1.0,"字符串转MD5失败")
        else:
            self.write_log_to_Text("ERROR:str_trans_to_md5 failed")


    #获取当前时间
    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        return current_time


    #日志动态打印
    def write_log_to_Text(self,logmsg):
        global LOG_LINE_NUM
        current_time = self.get_current_time()
        logmsg_in = str(current_time) +" " + str(logmsg) + "\n"      #换行
        if LOG_LINE_NUM <= 7:
            self.log_data_Text.insert(END, logmsg_in)
            LOG_LINE_NUM = LOG_LINE_NUM + 1
        else:
            self.log_data_Text.delete(1.0,2.0)
            self.log_data_Text.insert(END, logmsg_in)

    def ui_dir_tree_show(self, msg):
        self.dir_tree_text.delete('1.0', END)
        self.dir_tree_text.insert('1.0', msg)

    def keep_window(self):
        self.tk.mainloop()


ZMJ_PORTAL = MY_GUI()
def gui_start():
    # 设置根窗口默认属性
    ZMJ_PORTAL.set_init_window()
    ZMJ_PORTAL.keep_window()

def ui_log_show(msg):
    ZMJ_PORTAL.log_show(msg)
def ui_dir_tree_show(msg):
    ZMJ_PORTAL.ui_dir_tree_show(msg)
#    init_window.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


#if __name__ == '__main__':
def office_ui_main():
    def_log_module()
    gui_start()
