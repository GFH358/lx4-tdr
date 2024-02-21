#!/usr/bin/env python
#-*- coding:utf-8 -*-
from datetime import datetime, date, timedelta

import sum_table_import

import os, sys
import shutil
import zipfile
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

tdr_version = "v2.3"
developer_mode = 0       # 0：TDR模式；1：开发者模式
stable_en = 0


'''
注意事项：
    版本更新要更改标题栏和更新说明弹窗，关闭开发者模式！！！
    
使用规则：
1.外部信息读取兼容性较差，cfg文件修改系列和版本信息请务必按照原格式更改
    -系列与系列，版本与版本之间使用/隔开
    -在生成目录名时，系列信息只会提取数字部分用于组合，版本信息会直接使用用于组合
2.为保证结构完整，现版本不支持更改目录生成路径，SDK获取路径，所有资源默认与exe同一路径下
3.每次启动exe都会检查一遍目录结构，所以不用担心解压模板找不到位置，但是如果运行中删掉某些文件夹报错，那没办法了，懒得兼容
命名规则：
1.cfg-外部配置相关
2.sys-内部系统相关
3.app-软件前端相关
4.serial0：日常支持，serial1：测试协助
5.mode0：日常支持，mode1：测试协助
'''
cfg_data = []           # 存储外部配置
cfg_serial0 = ""        # 日常支持序号
cfg_serial1 = ""        # 测试协助序号
cfg_sdk_path = ""       # SDK获取地址，该功能暂不开放，默认路径在exe文件同级目录
cfg_target_path = ""    # 目录生成地址，该功能暂不开放，默认路径在exe文件同级目录
sys_time = ""           # init读取的系统时间
sys_theme = 0           # 系统主题 # 暂时不存在cfg中，会导致每次打开都是默认主题
sys_ui_reset_flag = 0   # UI复位标志
app_mode = "0"        # 模式0：普通模式，模式1：eso模式，内部模式，测试协助模式
# 系统存储芯片系列的列表，获取cfg信息后会把这里覆盖掉
sys_series_list = ['AC695N', 'AC696N', 'AC697N', 'AD698N', 'AC700N', 'JL701N', 'AC702N']
# 系统存储芯片系列对应SDK版本的列表，与sys_series_list对应
sys_version_list = [['100', '101', '102',],
                   ['200', '201', '202',],
                   ['300', '301', '302',],
                   ['400', '401', '402',],
                   ['500', '501', '502',],
                   ['600', '601', '602',],
                   ['700', '701', '702',],]

# 文件目录名配置
task_directory_path = "TaskDirectory"
task_record_path = "TaskRecord"
task_source_sdk_path = "TaskSourceSDK"
app_title = "Lairx4 Total Directory Report " + tdr_version  # 窗口标题

sys_cfg_path = "guser_config.txt"  # 存储外部配置的txt文件
sys_project_template_path = "ProjectTemplate.zip"  # 工程目录模板，必须包含PublicSDK文件夹用于解压SDK
sys_record_template_path = "RecordTemplate.md"      # 工程记录模板

# cfg文件解析配置
cfg_serial0_num_line = 0
cfg_serial1_num_line = 1
cfg_series_check_line = 2
cfg_series_list_line = 3
cfg_version_check_line = 4
cfg_version_list_line = 5
cfg_developer_tools_line = -1  # 规定开发者指令工具在cfg文件最后一行

# 开发者指令行解析 # 指令前有关键字【developer tools:】，所以有效位应该从16位后开始算
developer_mode_bit = 16             # 开发者模式:0关闭，1打开
developer_stable_en_bit = 17        # 总表自动生成:0关闭，1打开



# cfg文件解析校验配置
cfg_series_check_str = "series list:\n"
cfg_version_check_str = "version list:\n"
cfg_developer_tools_str = "developer tools:"     # 开发者指令关键字

# cfg文件不存在时，系统默认为cfg填充的值
sys_data = ["0\n",
            "0\n",
            cfg_series_check_str,
            "AC695N/AC696N/AC697N/AD698N/AC700N/JL701N/AC702N\n",
            cfg_version_check_str,
            "101/102/103\n",
            "101/102/103\n",
            "101/102/103\n",
            "101/102/103\n",
            "101/102/103\n",
            "101/102/103\n",
            "101/102/103\n",
            "注意：cfg文件数据读取没有做太多兼容和排错处理，请务按原格式修改本文件\n"
            "注意：cfg文件配置完成后，请在菜单栏设置read config或重启工具"]

# 用于检查GUI更新
old_app_series = ""
old_app_version = ""
old_app_scene = ""
old_app_production = ""
old_app_client = ""
old_app_title = ""
old_app_mode = ""
old_app_time = ""


'''********************************** 弹窗 *******************************************'''
# 首次使用弹窗内容
pop_explain_str = "\n\r" \
              "欢迎使用<Lx4-TDR " + tdr_version + ">\n\r" \
              "\n\r" \
              "基本文件目录已创建\n\r" \
              "guser_config.txt：请于此文件更改工具配置\n\r" \
              "TaskDirectory：任务目录生成存放于此\n\r" \
              "TaskRecord：任务记录生成存放于此\n\r" \
              "TaskSourceSDK：请于此添加源SDK压缩包，目前只支持zip格式\n\r" \
              "请于exe同级目录下更改工程模板<ProjectTemplate.zip>"
# v2.0更新说明
'''
pop_about_str = "Lairx4 Total Directory Report v2.0 \n\r" \
                "by flashy 2023.4.23\n\r" \
                "\n\r" \
                "Update description：\n\r" \
                "1.Add menu bar related functions\n\r" \
                "2.Add ESO mode\n\r" \
'''
# v2.1更新说明
'''
pop_about_str = "Lairx4 Total Directory Report " + tdr_version + " \n\r" \
                "by flashy 2023.7.23\n\r" \
                "\n\r" \
                "Update description：\n\r" \
                "1.Fix known issues\n\r" \
                "2.Add development manual\n\r" \
'''
# v2.2更新说明
'''
pop_about_str = "Lairx4 Total Directory Report " + tdr_version + " \n\r" \
                "by flashy 2023.7.23\n\r" \
                "\n\r" \
                "Update description：\n\r" \
                "1.Add sum table function\n\r" \
                "2.Fix known issues\n\r" \
'''
# v2.3更新说明
pop_about_str = "Lairx4 Total Directory Report " + tdr_version + " \n\r" \
                "by flashy 2024.3.1\n\r" \
                "\n\r" \
                "Update description：\n\r" \
                "1.Add developer tools\n\r" \
                "2.Fix known issues\n\r" \

# 帮助文档
pop_help_str = " ------------------------< 基本功能 >------------------------------\n\r" \
              "  此工具用于自动整合命名，快速生成任务目录和记录文档并解压SDK\n\r" \
              "\n\r" \
              " ----------------------< 基本目录结构 >----------------------------\n\r"  \
              "     guser_config.txt：请于此文件更改工具配置\n\r" \
              "     TaskDirectory：任务目录将存放于此\n\r" \
              "     TaskRecord：任务记录将存放于此\n\r" \
              "     TaskSourceSDK：请于此添加源SDK压缩包，目前只支持zip格式\n\r" \
              "\n\r" \
              " ----------------------< 模板更换说明 >----------------------------\n\r" \
              "     1. 工具在同级目录下以文件名读取模板，请务必保证模板名和路径无误\n\r" \
              "     2. 项目模板：ProjectTemplate.zip（需存在PublicSDK文件夹用于解压SDK）\n\r" \
              "     3. 记录模板：RecordTemplate.md\n\r" \
              "\n\r" \
              " ----------------------< 配置设置说明 >----------------------------\n\r" \
              "     1. 工具以行为单位读取config配置文件信息\n\r" \
              "     2. 第1、2行为工程生成序号（软件没兼容，请别瞎几把乱填非数字）\n\r" \
              "     3. 第3、5行为检验字符，请勿更改\n\r" \
              "     4. 第4行为芯片系列，以“ / ”字符隔开\n\r" \
              "     5. 第6至6+n行为版本系列，n为第4行芯片系列的个数，每行代表一个芯片系列的SDK版本\n\r" \
              "     6. 第6行对应第一个芯片的版本，第6+1行对应第二个芯片，第6+1行对应最后一个芯片，如此类推\n\r" \
              "\n\r" \
              " ------------------------<   其他   >------------------------------\n\r" \
              "     工具未作太多兼容，请尽可能不移动工具相关文件的路径\n\r" \
              "     获取SDK时会解压对应版本源文件夹内所有.zip文件\n\r" \
              "\n\r" \
              "终于写完了(～o￣3￣)～2023.4.22"

pop_gtools_help_str = " --------------------< 开发指令使用说明 >--------------------------\n\r" \
                 "     1. 开发者指令规定在cfg文件最末行，且有关键字[developer tools:]\n\r" \
                 "     2. 第一位开发者模式:0关闭，1打开\n\r" \
                 "     3. 第二位总表自动生成:0关闭，1打开\n\r" \
                 "     4. 其他位未定义"

developer_mode_bit = 16             # 开发者模式:0关闭，1打开
developer_stable_en_bit = 17        # 总表自动生成:0关闭，1打开
# 弹窗
def guser_popup_window(title_str, show_str):
    my_font = Font(family="宋体",size=12)
    tkinter.messagebox.showinfo(title_str, show_str)
'''********************************** 总表功能模块 *******************************************'''
import sum_table

stable_file_name ="任务目录汇总表.xlsx"
#stable_path = '..\\'                # 总表路径 (相对路径，上一级目录) 注意：sum_table.py也有该路径，需保持统一
stable_sheet_made0_name = "客户问题"        # 客户问题日常支持模式
stable_sheet_made1_name = "内部测试"        # 内部协助测试模式（ESO模式）
stable_add_list = ["TDR版本", "序号", "开始时间", "芯片系列", "SDK版本", "客户", "问题描述", "现场", "生产"]
# 用例：向stable_file_name文件的stable_sheet_made1_name表单最后一行写入stable_add_list列表的数据
# sum_table.stable_add_data(stable_add_list, stable_file_name,stable_sheet_made1_name)

# 获取将添加到总表的数据列表
def guser_get_add_stable_data(mode):
    global old_app_series
    global old_app_version
    global old_app_scene
    global old_app_production
    global old_app_client
    global old_app_title
    global old_app_mode
    global old_app_time
    global stable_add_list
    stable_add_list[0] = tdr_version
    if(mode == '0'):
        stable_add_list[1] = cfg_serial0
    elif(mode == '1'):
        stable_add_list[1] = "0x" + cfg_serial1
    stable_add_list[2] = old_app_time
    stable_add_list[3] = old_app_series
    stable_add_list[4] = old_app_version
    stable_add_list[5] = old_app_client
    stable_add_list[6] = old_app_title
    stable_add_list[7] = old_app_scene
    stable_add_list[8] = old_app_production
    return stable_add_list

# 向总表添加一行新数据
def guser_add_data_to_stable(mode,hyperlink):
    if stable_en == 0:
        return
    add_list = guser_get_add_stable_data(mode)      # 获写入总表的数据
    file_name = stable_file_name
    if(mode == '0'):
        sheet_made_name = stable_sheet_made0_name
    elif(mode == '1'):
        sheet_made_name = stable_sheet_made1_name
    if developer_mode == 1:
        sum_table.stable_add_data(add_list, file_name, sheet_made_name, hyperlink, "dev_mode")
        return
    else:
        sum_table.stable_add_data(add_list, file_name, sheet_made_name, hyperlink, "TDR_mode")
        return

'''********************************** 读写外部配置 *******************************************'''
# 从外部文件读取配置信息
def guser_read_cfg(path):
    if os.path.exists(path) == 0:        # 没有cfg文件，通常是第一次使用
        guser_popup_window("使用说明", pop_explain_str)             # 弹出使用说明
        guser_write_cfg(sys_data, path)  # 写cfg文件
        print("[warning] no cfg file！ New file successfully")
    with open(path, mode="r", encoding="utf-8") as f:
        data = f.readlines()  # read()    一次性读全部内容，以列表的形式返回结果。# readlines() 一行行读放在列表里
    f.close()
    print("sys read cfg data succeed\n")
    return data

# 向外部config文件写数据
def guser_write_cfg(datalist,savepath):
    del_datalist =['',]  # 写空列表删除数据
    for data in del_datalist:
        with open(savepath, mode="w", encoding="utf-8") as f:
            f.write(data)   # 写数据
    f.close()
    for data in datalist:
        with open(savepath, mode="a", encoding="utf-8") as f:
            f.write(data)       # 写数据
            # f.write("\n")     # 换行
    f.close()
    print("sys write cfg data succeed\n")

# 序号+1,传入参数：0是日常支持序号。1是测试协助序号
def guser_serial_add(mode):
    global cfg_serial0
    global cfg_serial1
    if mode == '0':
        output_data = int(cfg_serial0)
        cfg_serial0 = str(output_data + 1)
        print("cfg_serial0 +1 = %s" % cfg_serial0)
        return cfg_serial0
    elif mode == '1':
        output_data = int(cfg_serial1)
        cfg_serial1 = str(output_data + 1)
        print("cfg_serial1 +1 = %s" % cfg_serial1)
        return cfg_serial1

# 获取cfg配置的日常支持序号信息
def guser_get_serial0(cfg_data):
    if len(cfg_data) > cfg_serial0_num_line:
        serial0_int_data = int(cfg_data[cfg_serial0_num_line])
        if serial0_int_data >= 0 and serial0_int_data < 1000:
            output_serial0_data =str(serial0_int_data)
        else:
            output_serial0_data = "0"
            print("[warning] serial0 illegality, use -> sys_serial:0")
    else:
        output_serial0_data = "0"
        print("[warning] no cfg data, use -> sys_serial:0")
    print("serial0:", output_serial0_data)
    return output_serial0_data

# 获取cfg配置的测试协助序号信息
def guser_get_serial1(cfg_data):
    if len(cfg_data) > cfg_serial1_num_line:
        serial1_int_data = int(cfg_data[cfg_serial1_num_line])
        if serial1_int_data >= 0 and serial1_int_data < 1000:
            output_serial1_data =str(serial1_int_data)
        else:
            output_serial1_data = "0"
            print("[warning] serial1 illegality, use -> sys_serial:0")
    else:
        output_serial1_data = "0"
        print("[warning] no cfg data, use -> sys_serial:0")
    print("serial1:", output_serial1_data)
    return output_serial1_data

# 获取cfg配置的芯片系列信息
def guser_get_series_list(cfg_data):
    if len(cfg_data) > cfg_series_list_line and cfg_data[cfg_series_check_line] == cfg_series_check_str:
        data_str = cfg_data[cfg_series_list_line]
        output_series_list = [str(x) for x in data_str.strip().split("/")]
    else:  # 如果cfg文件没有数据，使用系统默认列表
        output_series_list = sys_series_list
        print("[warning] no cfg data, use -> sys_series_list")
    print("series list:", output_series_list)
    return output_series_list

# 获取cfg配置的SDK版本信息
def guser_get_version_list(cfg_data):
    long = len(sys_series_list)  # 依赖sys_series_list，需要结合guser_get_series_list()使用
    if len(cfg_data) > (cfg_version_check_line + long) and cfg_data[cfg_version_check_line] == cfg_version_check_str:
        valid_cfg_data = cfg_data[cfg_version_list_line: cfg_version_list_line + long]  # 在cfg数据中截取出version部分信息
        output_version_list = [[] for i in range(long)]  # 创建二维列表容器
        for i in range(long):
            output_version_list[i] = [str(x) for x in valid_cfg_data[i].strip().split("/")]
    else:
        output_version_list = sys_version_list
        print("[warning] no cfg data, use -> sys_version_list")
    print("version_list:", output_version_list)
    return output_version_list

# 获取开发者指令(gtools)
def guser_get_developer_tools(cfg_data):
    global developer_mode
    global stable_en
    if cfg_developer_tools_str in cfg_data[cfg_developer_tools_line]:  # 判断cfg文件末尾行是否有开发者指令关键字
        developer_order = cfg_data[cfg_developer_tools_line]
        developer_order_len = len(developer_order)

        if developer_order_len >= developer_mode_bit and developer_order[developer_mode_bit] == '0':
            developer_mode = 0
            print("[developer]:disable <developer_mode>!")
        elif developer_order_len >= developer_mode_bit and developer_order[developer_mode_bit] == '1':
            developer_mode = 1
            print("[developer]:enable <developer_mode>!")

        if developer_order_len >= developer_stable_en_bit and developer_order[developer_stable_en_bit] == '0':
            stable_en = 0
            print("[developer]:disable <stable_en>!")
        elif developer_order_len >= developer_stable_en_bit and developer_order[developer_stable_en_bit] == '1':
            stable_en = 1
            print("[developer]:enable <stable_en>!")

    else:
        print("[developer]:not developer_order!")


# 解包
def guser_cfg_decode(cfg_data):
    global cfg_serial0   # 声明cfg_serial为全局变量
    global cfg_serial1
    global sys_series_list
    global sys_version_list
    cfg_serial0 = guser_get_serial0(cfg_data)                       # 解析cfg数据获取序号信息 # 强转int再转str去掉\n
    cfg_serial1 = guser_get_serial1(cfg_data)                       # 解析cfg数据获取序号信息 # 强转int再转str去掉\n
    sys_series_list = guser_get_series_list(cfg_data)               # 解析cfg数据，获取系列信息
    sys_version_list = guser_get_version_list(cfg_data)             # 解析cfg数据，获取版本信息
    guser_get_developer_tools(cfg_data)                             # 开发者工具解析
    print("[decode]:decode cfg success")


'''********************************** SDK解压 *******************************************'''
# 解压zip
def guser_unzip_file(zip_file,target_dir):
    with zipfile.ZipFile(zip_file, "r") as zfile:
        for file in zfile.namelist():
            zfile.extract(file, target_dir)

# SDK拷贝器
def guser_sdk_copier(self,path):
    app_series = self.Combo1.get()
    app_version = self.Combo2.get()
    exist_zip_file = 0   # 存在zip文件的标志位
    # find_path：查找zip文件的文件目录；source_path：解压zip的源地址；target_path：zip解压目标地址
    find_path = task_source_sdk_path + '\\' + app_series + '\\' + app_version + '\\'
    if os.path.exists(find_path) != 0:
        file_list = os.listdir(find_path)
        print("file_list:", file_list)
        for i in range(len(file_list)):
            if ".zip" in file_list[i]:
                exist_zip_file = 1
                source_path = find_path + file_list[i]
                target_path = path + '\\PublicSDK'
                guser_unzip_file(source_path, target_path)
                print("copier:uncompress <%s> succeed" % file_list[i])
        if exist_zip_file == 0:
            print("[warning] copier:uncompress no SDK zip file!\n")
    else:
        print("[warning] copier:uncompress no SDK source directory!\n")


'''********************************** APP and Show *******************************************'''
# 获取eso模式开关状态
def guser_get_eso_switch(self):
    app_mode = self.Check3Var.get()
    return app_mode

# 获取GUI:芯片系列
def guser_get_chip_series(self):
    app_series = self.Combo1.get()
    return app_series

# 获取GUI:SDK版本
def guser_get_sdk_version(self):
    app_version = self.Combo2.get()
    return app_version

# 获取系统时间字符串
def guser_get_time():
    year = datetime.now().year
    month = datetime.now().month
    day = datetime.now().day
    if month < 10:
        str_month = '0' + str(month)
    else:
        str_month = str(month)
    if day < 10:
        str_day = '0' + str(day)
    else:
        str_day = str(day)
    str_time = str(year) + str_month + str_day
    return str_time

# 获取GUI:预览框输入字符串
def guser_get_preview_input(self):
    output_str = self.Text3.get()
    return output_str

# 获取预览框输出字符串
def guser_get_preview_output(self, mode):
    app_scene = self.Check1Var.get()
    app_production = self.Check2Var.get()
    app_client = self.Text1.get()
    app_title = self.Text2.get()
    app_time = sys_time
    app_series = guser_get_chip_series(self)
    use_series_str = ""         # 取出系列字符串中的数字部分用于组合目录文件名
    for i in range(len(app_series)):
        if app_series[i] >= "0" and app_series[i] <= "9":
            use_series_str = use_series_str + app_series[i]
    app_version = guser_get_sdk_version(self)

    if app_scene == '1':
        app_time = app_time + 'x'
    if app_production == '1':
        app_time = app_time + 's'
    # 日常支持
    if mode == '0':
        serial_num = int(cfg_serial0)
        if(serial_num >= 0 and serial_num <= 9):
            serial_num = str(serial_num)
            serial_num = '0' + str(serial_num)
        else:
            serial_num = cfg_serial0
        app_preview = serial_num + '.' + app_time + '_' + use_series_str + 'v' + app_version + '_' + app_client + '_' + app_title
    # 协助测试
    elif mode == '1':
        serial_num = int(cfg_serial1)
        if(serial_num >= 0 and serial_num <= 9):
            serial_num = str(serial_num)
            serial_num = '0' + str(serial_num)
        else:
            serial_num = cfg_serial1
        app_preview = '0x' + serial_num + '.' + app_time + '_' + use_series_str + 'v' + app_version + '_' + app_client + '_' + app_title
    else:
        app_preview = 'Error' + '.' + app_time + '_' + use_series_str + 'v' + app_version + '_' + app_client + '_' + app_title
        print("[preview]:mode error!")
    return app_preview

# 重构版本列表,用于美化前端版本号显示 #不使用，系统会获取框内数值，导致需要大改
def guser_create_display_version_list(version_list):
    print("guser_create_display_version_list")
    cnt = 0
    display_list = version_list
    for version_str in version_list:
        glen = len(version_str)
        output_str = ""
        for i in range(glen):
            if i == glen - 1:
                output_str = output_str + version_str[i]
            else:
                output_str = output_str + version_str[i] + '.'
        display_list[cnt] = 'v' + output_str
        cnt = cnt + 1
    print(display_list)
    return display_list


'''********************************** 菜单栏 *******************************************'''
# 菜单栏：打开
def guser_menu_open_directory():
    path = task_directory_path
    os.startfile(path)
    print("[menu]:open < %s > success" % path)
def guser_menu_open_record():
    path = task_record_path
    os.startfile(path)
    print("[menu]:open < %s > success" % path)
def guser_menu_open_source_sdk():
    path = task_source_sdk_path
    os.startfile(path)
    print("[menu]:open < %s > success" % path)
def guser_menu_open_sum_table():
    path = sum_table.stable_path + stable_file_name
    if os.path.exists(path) == 0:
        sum_table.stable_check_excel_exist(sum_table.test_excel_path) # 创建总表
    os.startfile(path)
    print("[menu]:open < %s > success" % stable_file_name)
def guser_menu_open_sum_table_import():
    top.destroy()
    sum_table_import.sti_app_start()
    print("[menu]:open < sum_table_import > success")

# 菜单栏：设置 open cfg
def guser_menu_setting():
    path = sys_cfg_path
    os.startfile(path)
    print("[menu]:open < %s > success" % path)

# 菜单栏：设置 read cfg
def guser_menu_read_cfg():
    global cfg_data
    global sys_ui_reset_flag
    cfg_data = guser_read_cfg(sys_cfg_path)
    guser_cfg_decode(cfg_data)
    sys_ui_reset_flag = 1
    guser_create_source_sdk()  # 10.19优化读取cfg文件后不更新SDK源目录问题
    print("[menu]:read < %s > success" % sys_cfg_path)
# 菜单栏：关于
def guser_menu_about():
    guser_popup_window("关于", pop_about_str)
    print("[menu]:pup < about > success")

# 菜单栏：帮助
def guser_menu_help():
    guser_popup_window("使用说明", pop_help_str)
    print("[menu]:pup < help > success")
    pass

# 菜单栏：gtools使用帮助
def guser_menu_gtools_help():
    guser_popup_window("开发指令使用说明", pop_gtools_help_str)
    print("[menu]:pup < gtools help > success")
    pass

# 菜单栏主题切换 #不加这么多花里胡哨的功能，不保存在cfg配置里，每次打开会恢复默认
def guser_menu_theme_switch():
    global sys_theme
    if sys_theme == 0:
        top.tk.call("set_theme", "light")
        # top.update()
    else:
        top.tk.call("set_theme", "dark")
        # top.update()
    sys_theme = (sys_theme + 1) % 2     # 取余，奇偶切换
    print("[menu]:theme_switch success")

# 创建菜单栏
def guser_app_create_menu():
    # 创建一个菜单
    menu = tkinter.Menu(top)
    # 创建子菜单
    filemenu = tkinter.Menu(menu, tearoff=0)
    filemenu.add_command(label=task_directory_path, command=guser_menu_open_directory)
    filemenu.add_command(label=task_record_path, command=guser_menu_open_record)
    filemenu.add_command(label=task_source_sdk_path, command=guser_menu_open_source_sdk)
    filemenu.add_separator()
    filemenu.add_command(label='sum_table', command=guser_menu_open_sum_table)
    filemenu.add_command(label='sum_table_import', command=guser_menu_open_sum_table_import)

    filemenu2 = tkinter.Menu(menu, tearoff=0)
    filemenu2.add_command(label="Open config", command=guser_menu_setting)
    filemenu2.add_command(label="Read config", command=guser_menu_read_cfg)

    filemenu3 = tkinter.Menu(menu, tearoff=0)
    filemenu3.add_command(label="tdr help", command=guser_menu_help)
    filemenu3.add_command(label="gtools help", command=guser_menu_gtools_help)
    # 将子菜单加入到菜单条中
    menu.add_cascade(label=u"打开", menu=filemenu)
    menu.add_cascade(label=u"设置", menu=filemenu2)
    menu.add_cascade(label=u"帮助", menu=filemenu3)

    # menu.add_command(label="帮助", command=guser_menu_help)
    menu.add_command(label="关于", command=guser_menu_about)

    menu.add_command(label="主题", command=guser_menu_theme_switch)
    # 添加到窗体中
    top.config(menu=menu)


'''********************************** 任务目录创建 *******************************************'''
# 创建SDK源目录
def guser_create_source_sdk():
    if os.path.exists(task_source_sdk_path) == 0:
        return
    else:
        for i in range(len(sys_series_list)):
            path = task_source_sdk_path + '\\' + sys_series_list[i]
            if os.path.exists(path) == 0:           # 检查目标文件夹是否已存在
                os.makedirs(path)                   # 生成系列文件夹
            # 在系列文件夹下再创建版本文件夹
            for j in range(len(sys_version_list[i])):
                path = task_source_sdk_path + '\\' + sys_series_list[i] + '\\' + sys_version_list[i][j]
                if os.path.exists(path) == 0:       # 检查目标文件夹是否已存在
                    os.makedirs(path)               # 生成版本文件夹

# 创建总工程目录
def guser_create_overall_directory():
    if os.path.exists(task_directory_path) == 0:    # 检查任务目录总文件夹
        os.makedirs(task_directory_path)            # 生成任务目录总文件夹
    if os.path.exists(task_record_path) == 0:       # 检查任务目录总文件夹
        os.makedirs(task_record_path)               # 生成任务记录总文件夹
    if os.path.exists(task_source_sdk_path) == 0:   # 检查任务目录总文件夹
        os.makedirs(task_source_sdk_path)           # 生成SDK源总文件夹
        guser_create_source_sdk()                   # 生成SDK源目录结构
    else:
        guser_create_source_sdk()

# 创建任务记录单项
def guser_create_record(app_preview):
    # 这里没兼容记录总目录被删或者新建md文件已存在的情况
    template_path = sys_record_template_path
    target = task_record_path + "\\" + app_preview + ".md"
    if os.path.exists(sys_record_template_path) == 0:  # 如果没有模板则生成一个
        print("[warning] no %s" % template_path)
        datastr = "请设置模板!\n"  # 写空列表删除数据
        with open(template_path, mode="w", encoding="utf-8") as f:
            f.write(datastr)  # 写数据
        f.close()

    shutil.copyfile(template_path, target)

# 创建任务目录单项
def guser_create_directory(self, app_preview):
    path = task_directory_path + '\\' + app_preview
    print("create directory:", path)
    cnt = 0
    while (os.path.exists(path)):  # 如果路径下该文件已经存在 # 有序号自加存在，基本不可能有这种情况
        cnt = cnt + 1
        scnt = str(cnt)
        path = task_directory_path + '\\' + app_preview + '(' + scnt + ')'
    os.makedirs(path) # 生成文件夹
    if os.path.exists(sys_project_template_path):  # 查找模板
        guser_unzip_file(sys_project_template_path, path)  # 解压工程模板
        if os.path.exists(path + '\\' + "PublicSDK") == 0:  # 查找模板
            os.makedirs(path + '\\' + "PublicSDK")  # 生成PublicSDK文件夹，防止模板中没PublicSDK文件夹
    else:  # 如果没有工程模板，直接生成文件夹
        os.makedirs(path + '\\' + "PublicSDK")  # 生成PublicSDK文件夹
        os.makedirs(path + '\\' + "ClientSDK")  # 生成ClientSDK文件夹
        print("[warning] no ProjectTemplate.zip")
    guser_sdk_copier(self, path)  # 向生成目录拷贝公版SDK

    os.startfile(path)  # 打开目录


'''********************************** 初始化 *******************************************'''
# 应用初始化
def guser_app_init():
    global sys_time
    global cfg_data
    sys_time = guser_get_time()                                     # 读取系统时间
    cfg_data = guser_read_cfg(sys_cfg_path)                         # 读取cfg数据存入系统
    guser_cfg_decode(cfg_data)
    guser_create_overall_directory()                                # 检查或创建总目录
    print('guser_app_init:ok')


class Application_ui(Frame):
    #这个类仅实现界面生成功能，具体事件处理代码在子类Application中。
    def __init__(self, master=None):
        Frame.__init__(self, master)

        self.master.title(app_title)
        self.master.geometry('604x369')
        self.master.resizable(0,0)          # 窗口大小不可调
        self.createWidgets()                # 创建部件

    '''********************************** 前端动态更新 *******************************************'''
# 更新前端初始化
    def update_check_init(self):
        global old_app_series
        global old_app_version
        global old_app_scene
        global old_app_production
        global old_app_client
        global old_app_title
        global old_app_mode
        global old_app_time
        old_app_series = self.Combo1.get()
        old_app_version = self.Combo2.get()
        old_app_scene = self.Check1Var.get()
        old_app_production = self.Check2Var.get()
        old_app_client = self.Text1.get()
        old_app_title = self.Text2.get()
        old_app_mode = guser_get_eso_switch(self)
        old_app_time = guser_get_time()
# 更新前端检查
    def update_check(self):
        global old_app_series
        global old_app_version
        global old_app_scene
        global old_app_production
        global old_app_client
        global old_app_title
        global old_app_mode
        global app_mode
        global old_app_time
        global sys_time
        app_series = self.Combo1.get()
        if app_series != old_app_series:
            old_app_series = app_series
            return 2
        app_mode = guser_get_eso_switch(self)
        if app_mode != old_app_mode:
            old_app_mode = app_mode
            return 1
        app_version = self.Combo2.get()
        if app_version != old_app_version:
            old_app_version = app_version
            return 1
        app_scene = self.Check1Var.get()
        if app_scene != old_app_scene:
            old_app_scene = app_scene
            return 1
        app_production = self.Check2Var.get()
        if app_production != old_app_production:
            old_app_production = app_production
            return 1
        app_client = self.Text1.get()
        if app_client != old_app_client:
            old_app_client = app_client
            return 1
        app_title = self.Text2.get()
        if app_title != old_app_title:
            old_app_title = app_title
            return 1
        sys_time = guser_get_time()
        if sys_time != old_app_time:
            old_app_time = sys_time
            return 1

        return 0

# 更新预览框
    def update_preview(self):
        self.Text3Var = StringVar(value=guser_get_preview_output(self, app_mode))
        self.Text3 = Entry(self.top, text='Text1', textvariable=self.Text3Var, font=('宋体', 12))
        self.Text3.place(relx=0.159, rely=0.65, relwidth=0.704, relheight=0.133)

#  读cfg文件后更新UI
    def reset_gui(self):
        # 系列下拉框
        self.Combo1List = sys_series_list
        self.Combo1 = Combobox(self.top, values=self.Combo1List, font=('宋体', 10))
        self.Combo1.place(relx=0.159, rely=0.385, relwidth=0.187, relheight=0.068)
        self.Combo1.set(self.Combo1List[0])
        # 版本下拉框
        self.Combo2List = sys_version_list[0]   # 默认显示第一个就行
        self.Combo2 = Combobox(self.top, values=self.Combo2List, font=('宋体', 10))
        self.Combo2.place(relx=0.437, rely=0.385, relwidth=0.187, relheight=0.068)
        self.Combo2.set(self.Combo2List[0])
        # 预览显示框
        self.Text3Var = StringVar(value=guser_get_preview_output(self, app_mode))
        self.Text3 = Entry(self.top, text='Text1', textvariable=self.Text3Var, font=('宋体', 12))
        self.Text3.place(relx=0.159, rely=0.65, relwidth=0.704, relheight=0.133)

# 动态更新前端
    def update(self):
        global sys_ui_reset_flag
        if sys_ui_reset_flag == 1:
            sys_ui_reset_flag = 0
            self.reset_gui()
        update_check_flag = self.update_check()  # 获取GUI数据更新
        if update_check_flag == 0:
            self.Combo2.after(500, self.update)  # 无更新，重设定时器
            # print(".") # 开这个打印检查exe是否卡死
            return
        else:
            print("GUI:change")
            if update_check_flag == 2:
                lcnt = 0
                for i in sys_series_list:
                    if self.Combo1.get() == i:
                        self.Combo2List = sys_version_list[lcnt]
                        break
                    else:
                        lcnt = lcnt + 1
                # 更新版本下拉框
                self.Combo2 = Combobox(self.top, values=self.Combo2List, font=('宋体', 10))
                self.Combo2.place(relx=0.437, rely=0.385, relwidth=0.187, relheight=0.068)
                self.Combo2.set(self.Combo2List[0])
            # 更新预览框
            self.update_preview()
            # 完成更新，重设定时器
            self.Combo2.after(500, self.update)

    '''********************************** GUI生成 *******************************************'''
    def createWidgets(self):
        self.top = self.winfo_toplevel()  # 获取顶层
        self.style = Style()
# 现场勾选框
        self.Check1Var = StringVar(value='0')
        self.style.configure('Check1.TCheckbutton', font=('宋体',9))
        self.Check1 = Checkbutton(self.top, text='现场', variable=self.Check1Var, style='Check1.TCheckbutton')
        self.Check1.place(relx=0.636, rely=0.173, relwidth=0.134, relheight=0.08)
# 生产勾选框
        self.Check2Var = StringVar(value='0')
        self.style.configure('Check2.TCheckbutton', font=('宋体',9))
        self.Check2 = Checkbutton(self.top, text='生产', variable=self.Check2Var, style='Check2.TCheckbutton')
        self.Check2.place(relx=0.750, rely=0.173, relwidth=0.134, relheight=0.08)
# switch开关
        self.Check3Var = StringVar(value='0')
        self.style.configure('Check2.TCheckbutton', font=('宋体', 9))
        self.Check3 = Checkbutton(self.top, text='ESO', variable=self.Check3Var, style='Switch.TCheckbutton')
        self.Check3.place(relx=0.715, rely=0.385, relwidth=0.150, relheight=0.078)
# 客户输入框
        self.Text1Var = StringVar(value='客户')
        self.Text1 = Entry(self.top, text='Text1', textvariable=self.Text1Var, font=('宋体',9))
        self.Text1.place(relx=0.159, rely=0.173, relwidth=0.439, relheight=0.08)
# 标题输入框
        self.Text2Var = StringVar(value='标题')
        self.Text2 = Entry(self.top, text='Text1', textvariable=self.Text2Var, font=('宋体', 9))
        self.Text2.place(relx=0.159, rely=0.282, relwidth=0.704, relheight=0.08)
# 系列下拉框
        self.Combo1List = sys_series_list
        self.Combo1 = Combobox(self.top, values=self.Combo1List, font=('宋体', 10))
        self.Combo1.place(relx=0.159, rely=0.385, relwidth=0.187, relheight=0.068)
        self.Combo1.set(self.Combo1List[0])
# 版本下拉框
        lcnt = 0
        for i in sys_series_list:
            if self.Combo1.get() == i:
                self.Combo2List = sys_version_list[lcnt]
                break
            else:
                lcnt = lcnt + 1
        self.Combo2 = Combobox(self.top, values=self.Combo2List, font=('宋体', 10))
        self.Combo2.place(relx=0.437, rely=0.385, relwidth=0.187, relheight=0.068)
        self.Combo2.set(self.Combo2List[0])


# 预览显示框
        self.Text3Var = StringVar(value=guser_get_preview_output(self, app_mode))
        self.Text3 = Entry(self.top, text='Text1', textvariable=self.Text3Var, font=('宋体', 12))
        self.Text3.place(relx=0.159, rely=0.65, relwidth=0.704, relheight=0.133)
# make按钮
        self.style.configure('Command1.TButton',font=('宋体',9))
        self.Command1 = Button(self.top, text='M A K E', command=self.Command1_Cmd, style='Command1.TButton')  # 绑定回调Command1_Cmd
        self.Command1.place(relx=0.384, rely=0.867, relwidth=0.161, relheight=0.111)
# 文本
        self.style.configure('Label2.TLabel',anchor='w', font=('宋体',9))
        self.Label2 = Label(self.top, text='预览', style='Label2.TLabel')
        self.Label2.place(relx=0.079, rely=0.684, relwidth=0.055, relheight=0.068)

        self.style.configure('Label2.TLabel',anchor='w', font=('宋体',9))
        self.Label2 = Label(self.top, text='客户', style='Label2.TLabel')
        self.Label2.place(relx=0.079, rely=0.173, relwidth=0.055, relheight=0.068)

        self.style.configure('Label2.TLabel',anchor='w', font=('宋体',9))
        self.Label2 = Label(self.top, text='版本', style='Label2.TLabel')
        self.Label2.place(relx=0.371, rely=0.385, relwidth=0.055, relheight=0.068)

        self.style.configure('Label2.TLabel',anchor='w', font=('宋体',9))
        self.Label2 = Label(self.top, text='芯片', style='Label2.TLabel')
        self.Label2.place(relx=0.079, rely=0.385, relwidth=0.055, relheight=0.068)

        self.style.configure('Label2.TLabel',anchor='w', font=('宋体',9))
        self.Label2 = Label(self.top, text='标题', style='Label2.TLabel')
        self.Label2.place(relx=0.079, rely=0.282, relwidth=0.055, relheight=0.068)

        self.update_check_init()             # UI更新初始化
        self.Combo2.after(500, self.update)  # 设置UI实时更新定时器
        # guser_guide_window()


# MAKE按键回调实现！！！！！！！！！
class Application(Application_ui):
    # 这个类实现具体的事件处理回调函数。界面生成代码在Application_ui中。
    def __init__(self, master=None):
        Application_ui.__init__(self, master)

    def Command1_Cmd(self, event=None):
        #TODO, Please finish the function here!
        #2.18优化
        global cfg_data
        cfg_data = guser_read_cfg(sys_cfg_path)  # 读取cfg数据存入系统
        guser_cfg_decode(cfg_data)  # 解析cfg文件
        self.update_preview()  # 更新预览框

        app_preview = guser_get_preview_input(self)  # 获取当前GUI设置

        guser_add_data_to_stable(app_mode, app_preview)  # 添加总表

        if app_mode == '0':
            guser_create_directory(self, app_preview)  # 生成目录
            guser_create_record(app_preview)  # 生成记录md文件
            cfg_data[cfg_serial0_num_line] = guser_serial_add(app_mode) + '\n'
            print("[MAKE]:create directory and record success!")
        elif app_mode == '1':
            guser_create_directory(self, app_preview)  # eso模式只生成目录，不生成记录md文件
            cfg_data[cfg_serial1_num_line] = guser_serial_add(app_mode) + '\n'
            print("[MAKE]:ESO mode create directory success!")
        guser_write_cfg(cfg_data, sys_cfg_path)  # 写cfg文件
        self.update_preview()
        top.state("iconic")  # 窗口有3中状态，iconic：最小化；normal：正常显示；zoomed：最大化


def testfun():
    print("test ok")
# task_directory_path = "TaskDirectory"
# task_record_path = "TaskRecord"
# task_source_sdk_path = "TaskSourceSDK"

if __name__ == "__main__":
    top = Tk()                                  # 实例化一个Tk窗口对象
    guser_app_create_menu()                     # 创建菜单栏

    # top.wm_attributes("-topmost", True)       # 设置GUI置顶
    top.wm_attributes("-alpha", 0.85)           # 设置GUI透明度(0.0~1.0)
    # top.attributes("-alpha", 0.9)             # 设置GUI透明度(0.0~1.0)

    si = Sizegrip(top, style='1.TSizegrip')     # 右下角定位三角
    si.pack(side=BOTTOM, anchor=SE)

    top.tk.call("source", "azure.tcl")          # 导入主题资源
    top.tk.call("set_theme", "dark")            # 默认主题：dark
    # top.tk.call("set_theme", "light")         # 默认主题：light
    # top.iconbitmap(icon_file) # icon_file就是一个.ico的图标文件，使用绝对或相对路径
    guser_app_init()                            # 应用初始化
    Application(top).mainloop()                 # mainloop()显示窗口
    try: top.destroy()
    except: pass
