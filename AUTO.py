# -*- coding: UTF-8 -*-
import snap7
import pandas as pd
from snap7.util import get_bool
from snap7.util import get_real  # 仅导入需要的函数
from snap7.util import get_string
from datetime import datetime
from time import sleep
import os

import sys
import importlib

_orig_import = importlib.import_module

def _tracked_import(name, *args, **kwargs):
    print(f"Importing: {name}")  # 输出到控制台或日志文件
    return _orig_import(name, *args, **kwargs)

importlib.import_module = _tracked_import

# 你的主程序代码...
#os.system("reg add HKEY_CURRENT_USER\Console /v QuickEdit /t REG_DWORD /d 0 /f")

def bool_trigger_callback(new_value):
    """布尔值变化时的回调函数"""
    if new_value:
        data1=read_plc_real(plc, DB_NUMBER, BYTE_INDEX)
        time_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # 添加时间
        data2 = read_plc_string(plc, DB_NUMBER, 10, 12)
        data3=read_plc_bool(plc, DB_NUMBER, 0,4)
        data={"圆度检测值": data1,'检测时间':time_now,'工单号': data2,'检测合格':data3}
        try:
            df_existing = pd.read_excel("检测数据.xlsx")
            df_new = pd.DataFrame([data])
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        except FileNotFoundError:
            df_combined = pd.DataFrame([data])

        # 使用 xlsxwriter 引擎美化
        writer = pd.ExcelWriter("检测数据.xlsx", engine="xlsxwriter")
        df_combined.to_excel(writer, index=False, sheet_name="Sheet1")

        # 获取 workbook 和 worksheet 对象
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        # 定义格式
        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#4F81BD",
            "font_color": "white",
            "border": 1,
            "align": "center"
        })

        data_format = workbook.add_format({
            "border": 1,
            "align": "center"
        })

        # 应用格式
        for col_num, value in enumerate(df_combined.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # 设置列宽
        worksheet.set_column(0, len(df_combined.columns) - 1, 20, data_format)  # 统一列宽为20
        # 保存
        writer.close()
        print("数据已成功写入 检测数据.xlsx")

    else:
        print("等待PLC数据...")


def read_plc_bool(plc, db_number, byte_offset, bit_offset):
    """读取PLC中指定布尔值"""
    data = plc.db_read(db_number, byte_offset, 1)  # 读取1个字节
    byte_value = data[0]  # 提取字节的整数值
    bool_value = (byte_value & (1 << bit_offset)) != 0  # 提取指定位的值
    return bool_value

def read_plc_real(plc, db_number, byte_index):
    """读取PLC中指定布尔值"""
    data_1 = plc.db_read(db_number, byte_index, 4)
    real_value = get_real(data_1, 0)
    return real_value

def read_plc_string(plc, db_number, byte_offset_1,byte_index_1):
    data_2 = plc.db_read(db_number, byte_offset_1,byte_index_1)
    string_value = get_string(data_2, 0)
    return string_value

# 连接参数
PLC_IP = '192.168.1.100'
RACK = 0
SLOT = 1
DB_NUMBER = 24
BYTE_OFFSET = 0  # 假设布尔变量在DB20.DBB0.0
BIT_OFFSET = 0
BYTE_INDEX = 2
BYTE_OFFSET_1 = 10

# 初始化PLC连接
plc = snap7.client.Client()

plc.connect(PLC_IP, RACK, SLOT)

try:
    last_value = None
    while True:
        # 读取当前布尔值
        current_value = read_plc_bool(plc, DB_NUMBER, BYTE_OFFSET, BIT_OFFSET)

        # 检查值是否变化
       # if last_value is not None and current_value != last_value:
        if last_value is False and current_value is True:
            bool_trigger_callback(current_value)  # 触发回调函数

        last_value = current_value  # 更新上一次的值
        sleep(1)  # 控制轮询间隔（单位：秒）

except KeyboardInterrupt:
    print("用户中断监控")
finally:
    plc.disconnect()
