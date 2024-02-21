#!/usr/bin/env python
#-*- coding:utf-8 -*-

import os, sys
try:
    from tkinter import *
except ImportError:  #Python 2.x
    PythonVersion = 2
    from Tkinter import *
    from tkFont import Font
    from ttk import *
    #Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel
    from tkMessageBox import *
    #Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')
    #import tkFileDialog
    #import tkSimpleDialog
else:  #Python 3.x
    PythonVersion = 3
    from tkinter.font import Font
    from tkinter.ttk import *
    from tkinter.messagebox import *
    #import tkinter.filedialog as tkFileDialog
    #import tkinter.simpledialog as tkSimpleDialog    #askstring()

import tkinter
import tkinter.messagebox
from tkinter import filedialog

import sum_table

'''
命名规则：
1.sti-总表功能相关
2.app-软件前端相关
3.tdr-TDR导入相关
'''


sti_input_version = ""                   # 用户输入的TDR版本号
sti_input_path = ""                      # 用户输入的TDR目录路径
tdr_directory_name = "TaskDirectory"      # TDR任务目录文件夹
tdr_record_name = "TaskRecord"            # TDR任务记录文件夹
tdr_directory_exist = 0                  # TDR存在目录标志
tdr_record_exist = 0                     # TDR存在记录标志
sti_stable_path = ""                     # 总表路径
sti_stable_name = "任务目录汇总表.xlsx"                     # 总表文件名
sti_sheet_made0_name = "客户问题"        # 客户问题日常支持模式
sti_sheet_made1_name = "内部测试"        # 内部协助测试模式（ESO模式）

app_default_version_text = "    请输入TDR版本号"
app_default_path_text = "    请输入TDR目录路径"

app_version_errpop_title = "出错啦~"
app_version_errpop_show = "请输入TDR版本号"
app_path_errpop_title = "出错啦~"
app_path_errpop_show = "请检查TDR路径是否有效"

import_update_list = ["导入中，请勿关闭STI.",
                      "导入中，请勿关闭STI. .",
                      "导入中，请勿关闭STI. . ."]
import_update_cnt = 0
import_data_cnt = 0
import_is_run = 0

'''********************************** 总表导入 *******************************************'''
# 解析出错
def sti_tdr_resolver_err(str):
    print("[error] -> [STI]:resolver file err : %s\n" % str)

# TDR解析器
def sti_tdr_resolver(input_str):
    # 输入str，拆分各个excel单元格单项,返回一个列表
    counter = 0
    input_str_len = len(input_str)
    output_serial = ""       # 序号
    output_time = ""         # 开始时间
    output_series = ""       # 芯片系列
    output_version = ""      # SDK版本
    output_client = ""       # 客户
    output_title = ""        # 问题
    output_scene = "N"       # 现场
    output_production = "N"  # 生产

    output_list = ["TDR版本", "序号", "开始时间", "芯片系列", "SDK版本", "客户", "问题描述", "现场", "生产"]

    # 获取序号
    if input_str[:2] == "0x":   # 内部模式
        output_serial = "0x"
        counter = 2
    for i in range(counter, input_str_len):
        if input_str[i] >= '0' and input_str[i] <= '9':
            output_serial = output_serial + input_str[i]
        elif input_str[i] == '.':
            counter = i + 1
            break
        else:
            sti_tdr_resolver_err(input_str)
            return 0
    # 获取时间，现场，生产
    for i in range(counter, input_str_len):
        if input_str[i] >= '0' and input_str[i] <= '9':
            output_time = output_time + input_str[i]
        elif input_str[i] == 'x':
            output_scene = 'Y'
        elif input_str[i] == 's':
            output_production = 'Y'
        elif input_str[i] == '_':
            counter = i + 1
            break
        else:
            sti_tdr_resolver_err(input_str)
            return 0
    # 获取芯片系列、SDK版本
    version_flag = 0
    for i in range(counter, input_str_len):
        if input_str[i] == 'v':
            version_flag = 1
        elif input_str[i] != '_':
            if version_flag == 0:
                output_series = output_series + input_str[i]
            else:
                output_version = output_version + input_str[i]
        else:
            counter = i + 1
            break
        if i == input_str_len:
            sti_tdr_resolver_err(input_str)
            return 0
    # 获取客户
    for i in range(counter, input_str_len):
        if input_str[i] != '_':
            output_client = output_client + input_str[i]
        else:
            counter = i + 1
            break
        if i == input_str_len:
            sti_tdr_resolver_err(input_str)
            return 0
    # 获取任务标题
    for i in range(counter, input_str_len):
        output_title = output_title + input_str[i]
    if ".md" in output_title:
        output_title = output_title[:-3]

    output_list[0] = sti_input_version
    output_list[1] = output_serial
    output_list[2] = output_time
    output_list[3] = output_series
    output_list[4] = output_version
    output_list[5] = output_client
    output_list[6] = output_title
    output_list[7] = output_scene
    output_list[8] = output_production
    # print(output_list)
    return output_list



# Excel导入器
def sti_import_excel(import_data, mode, hyperlink):
    # 向总表添加一行新数据，模式：0日常支持，1内部测试
    add_list = import_data  # 获写入总表的数据
    file_name = sti_stable_name
    if (mode == '0'):
        sheet_made_name = sti_sheet_made0_name
    elif (mode == '1'):
        sheet_made_name = sti_sheet_made1_name
    print("import:%s" % hyperlink)
    hyperlink_path = sti_input_path
    sum_table.stable_add_data(add_list, file_name, sheet_made_name, hyperlink, hyperlink_path)
    return


def sti_popup_window(title_str, show_str):
    global import_is_run
    import_is_run = 0
    my_font = Font(family="宋体",size=12)
    tkinter.messagebox.showinfo(title_str, show_str)



class Application2_ui(Frame):
    #这个类仅实现界面生成功能，具体事件处理代码在子类Application2中。
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('Sum Table Import')
        self.master.geometry('451x163')
        self.createWidgets()

    # 总表导入启动
    # 1.有效性检测，把TDR文件夹path获取下来
    # 2.获取TDR列表
    # 3.解析器：解析TDR单项
    # 4.导入器：写入excel总表
    # 5.更新UI
    # 6.TDR列表导入未结束，回到3循环
    def app_sti_import_start(self):
        global sti_input_version
        global sti_input_path
        global tdr_directory_exist
        global tdr_record_exist
        global import_data_cnt
        global import_is_run
        if import_is_run == 1:  # 若STI正在导入中，按import无效
            print("[STI]:sti is busy, import key close\n")
            return
        else:
            import_is_run = 1
        # 检查用户输入信息
        sti_input_version = self.entry1Var.get()                # 获取客户输入的TDR版本
        sti_input_path = self.entry2Var.get()                   # 获取客户输入的TDR路径
        if sti_input_version == "":                             # 检查用户输入TDR版本号是否为空
            print("[error] -> [STI]:TDR version is null\n")
            sti_popup_window(app_version_errpop_title, app_version_errpop_show)
            return
        elif sti_input_path == "":                              # 检查用户输入TDR路径是否为空
            print("[error] -> [STI]:TDR path is null\n")
            sti_popup_window(app_path_errpop_title, app_path_errpop_show)
            return

        # 检查TDR路径有效性，directory和record至少存在一个
        tdr_directory_path = sti_input_path + '\\' + tdr_directory_name
        tdr_record_path = sti_input_path  + '\\' + tdr_record_name
        if os.path.exists(tdr_directory_path) == 1:         # 检查TDR目录文件夹是否已存在
            tdr_directory_exist = 1
            print("[STI]:TDR directory path exist\n")
        if os.path.exists(tdr_record_path) == 1:            # 检查TDR记录文件夹是否已存在
            tdr_record_exist = 1
            print("[STI]:TDR record path exist\n")
        if tdr_directory_exist or tdr_record_exist:         # 目录一个都不存在，路径无效
            print("[STI]:sti start success\n")
        else:
            print("[error] -> [STI]:TDR path is invalid\n")
            sti_popup_window(app_path_errpop_title, app_path_errpop_show)
            return

        # 获取列表
        import_data_cnt = 0
        if tdr_directory_exist == 1:  # directory目录存在
            directory_list = os.listdir(tdr_directory_path)
            for i in range(len(directory_list)):
                import_data_list = sti_tdr_resolver(directory_list[i])
                if import_data_list == 0:  # 解析文件失败，跳过本次循环
                    continue
        # 写入excel
                import_data_cnt = import_data_cnt+1
                self.app_import_update_entry(directory_list[i])
                if "0x" in import_data_list[1]:
                    sti_import_excel(import_data_list, '1', directory_list[i])
                else:
                    sti_import_excel(import_data_list, '0', directory_list[i])
        elif tdr_record_exist == 1:  # directory目录不存在，record目录存在
            print("[STI]:not [directory] file, use [record] file\n")
            record_list = os.listdir(tdr_record_path)
            for i in range(len(record_list)):
                import_data_list = sti_tdr_resolver(record_list[i])
                if import_data_list == 0:  # 解析文件失败，跳过本次循环
                    continue
        # 写入excel
                record_list[i] = record_list[i][:-3]  # 去掉.md字符
                import_data_cnt = import_data_cnt + 1
                self.app_import_update_entry(record_list[i])
                if "0x" in import_data_list[1]:
                    sti_import_excel(import_data_list, '1', record_list[i])
                else:
                    sti_import_excel(import_data_list, '0', record_list[i])
        import_is_run = 0
        self.app_update_success_entry(import_data_cnt)  # 更新导入完成ui

        return

    # 导入中更新输入框文本
    def app_import_update_entry(self, show_str):
        global import_update_cnt
        self.entry1Var.set(import_update_list[import_update_cnt])
        self.entry_name1.update()   # 更新界面
        import_update_cnt = import_update_cnt + 1
        if import_update_cnt > len(import_update_list) - 1:
            import_update_cnt = 0
        self.entry2Var.set(show_str)
        self.entry_name2.update()
        return

    # 导入完成更新输入框文本
    def app_update_success_entry(self, import_data_cnt):
        output_str = "导入成功，导入数据：" + str(import_data_cnt)
        self.entry1Var.set("")
        self.entry2Var.set(output_str)
        self.entry_name1.update()
        self.entry_name2.update()

    # 更新路径输入框文本
    def app_update_path_entry(self):
        self.app_entry_clear_text()
        path = filedialog.askdirectory()
        self.entry2Var.set(path)

    # 清除输入框默认文本并恢复字体颜色
    def app_entry_clear_text(self):
        print("sti_entry_clear_text")
        if self.entry1Var.get() == app_default_version_text:
            self.entry1Var.set('')
            self.entry_name1 = Entry(self.sti_top, textvariable=self.entry1Var, width=35)  # 恢复默认黑色字体颜色，关闭焦点检测清除
            self.entry_name1.place(relx=0.215, rely=0.213)
            # self.entry_name1.delete(0, 'end') # 清除Entry文本
        if self.entry2Var.get() ==app_default_path_text:
            self.entry2Var.set('')
            self.entry_name2 = Entry(self.sti_top, textvariable=self.entry2Var, width=35)
            self.entry_name2.place(relx=0.215, rely=0.442)

    def createWidgets(self):
        self.sti_top = self.winfo_toplevel()

        self.style = Style()

        # 文本
        self.style.configure('Label1.TLabel',anchor='w', font=('宋体',9))
        self.Label1 = Label(self.sti_top, text='TDR VERSION', style='Label1.TLabel')
        self.Label1.place(relx=0.050, rely=0.210, relwidth=0.18, relheight=0.153)
        self.style.configure('Label2.TLabel',anchor='w', font=('宋体',9))
        self.Label2 = Label(self.sti_top, text='TDR PATH', style='Label2.TLabel')
        self.Label2.place(relx=0.050, rely=0.442, relwidth=0.18, relheight=0.153)

        # 输入框
        self.entry1Var = StringVar(value = app_default_version_text)
        self.entry_name1 = Entry(self.sti_top, textvariable=self.entry1Var, width=35, foreground='gray', validate='focusin', validatecommand=self.app_entry_clear_text)
        self.entry_name1.place(relx=0.215, rely=0.213)
        self.entry2Var = StringVar(value = app_default_path_text)  # 文件输入路径变量
        self.entry_name2 = Entry(self.sti_top, textvariable=self.entry2Var, width=35, foreground='gray', validate='focusin', validatecommand=self.app_entry_clear_text)
        self.entry_name2.place(relx=0.215, rely=0.442)

        # 按键
        self.style.configure('Command2.TButton', font=('宋体', 9))
        self.Command2 = Button(self.sti_top, text="路径选择", command=self.app_update_path_entry)
        self.Command2.place(relx=0.788, rely=0.200)
        self.style.configure('Command2.TButton', font=('宋体', 9))
        self.Command2 = Button(self.sti_top, text="路径选择", command=self.app_update_path_entry)
        self.Command2.place(relx=0.788, rely=0.426)

        self.style.configure('Command1.TButton',font=('宋体',9))
        self.Command1 = Button(self.sti_top, text='I M P O R T', command=self.app_sti_import_start, style='Command1.TButton')
        self.Command1.place(relx=0.355, rely=0.687, relwidth=0.300, relheight=0.202)


class Application2(Application2_ui):
    #这个类实现具体的事件处理回调函数。界面生成代码在Application2_ui中。
    def __init__(self, master=None):
        Application2_ui.__init__(self, master)

    def Command1_Cmd(self, event=None):
        #TODO, Please finish the function here!
        pass

def sti_app_start():
    sti_top = Tk()
    Application2(sti_top).mainloop()
    try: sti_top.destroy()
    except: pass

# sti_app_start()