"""
配置文件
"""

import os
import sys

# 获取应用根目录（开发环境使用脚本所在目录，打包环境使用exe所在目录）
if getattr(sys, 'frozen', False):
    # PyInstaller打包后的exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # 开发环境
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 数据文件路径
DATA_DIR = os.path.join(BASE_DIR, "data")
TEACHERS_FILE = os.path.join(DATA_DIR, "teachers.xlsx")
EXAMS_FILE = os.path.join(DATA_DIR, "exams.xlsx")
SCHEDULE_FILE = os.path.join(DATA_DIR, "schedule.xlsx")
CONFIG_FILE = os.path.join(DATA_DIR, "config.xlsx")

# 默认排班配置
DEFAULT_MAX_EXAMS_PER_DAY = 3
DEFAULT_MAX_CONSECUTIVE_EXAMS = 2
DEFAULT_TIME_SLOTS = ["08:30-10:30", "10:45-12:45", "14:00-16:00", "16:15-18:15"]

# Excel 导出配置
EXPORT_ENCODING = "utf-8-sig"
