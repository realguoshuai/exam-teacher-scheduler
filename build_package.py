#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
监考老师排班系统 - 打包脚本
"""

import os
import sys
import shutil
import subprocess
from datetime import datetime

def print_step(step_num, total_steps, message):
    print(f"[{step_num}/{total_steps}] {message}")

def run_command(cmd, step_info):
    print(f"  执行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
    if result.returncode != 0:
        print(f"  {step_info} 失败")
        print(f"  错误: {result.stderr}")
        return False
    return True

def build():
    print("="*50)
    print("监考老师排班系统 - 打包工具")
    print("="*50)
    print()

    total_steps = 6

    # 步骤1: 检查Python
    print_step(1, total_steps, "检查Python环境")
    if sys.version_info < (3, 6):
        print("  错误: 需要Python 3.6或更高版本")
        print(f"  当前版本: {sys.version}")
        return False
    print(f"  Python版本: {sys.version}")
    print()

    # 步骤2: 安装依赖
    print_step(2, total_steps, "安装打包依赖")
    packages = ['pyinstaller', 'flask', 'flask-cors', 'pandas', 'openpyxl', 'matplotlib']
    for pkg in packages:
        cmd = [sys.executable, '-m', 'pip', 'install', pkg, '-i', 'https://pypi.tuna.tsinghua.edu.cn/simple']
        result = subprocess.run(cmd, capture_output=True)
        if result.returncode != 0:
            print(f"  安装 {pkg} 失败")
            return False
        print(f"  {pkg} OK")
    print()

    # 步骤3: 清理旧文件
    print_step(3, total_steps, "清理旧的打包文件")
    dirs_to_clean = ['dist_release', 'build', '.venv_package']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"  已删除: {dir_name}")
    print()

    # 步骤4: 执行PyInstaller打包
    print_step(4, total_steps, "执行PyInstaller打包")
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--clean',
        '--onefile',
        '--name', '监考老师排班系统',
        '--add-data', 'templates;templates',
        '--hidden-import=flask',
        '--hidden-import=openpyxl',
        '--hidden-import=openpyxl.drawing.image',
        '--hidden-import=matplotlib',
        '--hidden-import=pandas',
        '--distpath', 'dist_release',
        'app.py'
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print("  打包失败!")
        print(result.stderr)
        return False
    print("  打包成功")
    print()

    # 步骤5: 复制必要文件
    print_step(5, total_steps, "复制必要文件")
    output_dir = 'dist_release'

    # 复制templates
    if os.path.exists('templates'):
        dest_templates = os.path.join(output_dir, 'templates')
        if os.path.exists(dest_templates):
            shutil.rmtree(dest_templates)
        shutil.copytree('templates', dest_templates)
        print("  templates 已复制")

    # 创建data目录
    data_dir = os.path.join(output_dir, 'data')
    os.makedirs(data_dir, exist_ok=True)
    print("  data 目录已创建")

    # 复制启动脚本和文档
    files_to_copy = [
        '使用文档.md'
    ]
    for filename in files_to_copy:
        if os.path.exists(filename):
            shutil.copy2(filename, os.path.join(output_dir, filename))
            print(f"  {filename} 已复制")

    # 创建使用说明
    readme_content = """监考老师排班系统
=================

[使用方法]
方法1（推荐）： 双击运行 "启动程序.bat"
方法2： 直接双击运行 "监考老师排班系统.exe"

[注意事项]
- 首次运行会自动创建 data 目录和示例数据
- 请确保端口 5000 没有被其他程序占用
- 关闭程序时，请直接关闭命令行窗口

版本：1.0
打包日期：{}
""".format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

    with open(os.path.join(output_dir, '使用说明.txt'), 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print("  使用说明.txt 已创建")
    print()

    # 步骤6: 完成
    print_step(6, total_steps, "打包完成!")
    exe_path = os.path.join(output_dir, '监考老师排班系统.exe')
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print(f"  可执行文件: {exe_path}")
        print(f"  文件大小: {size_mb:.1f} MB")
    print()
    print("="*50)
    print("可以将 dist_release 文件夹发送给其他电脑使用")
    print("="*50)

    return True

if __name__ == '__main__':
    success = build()
    input("\n按回车键退出...")
    sys.exit(0 if success else 1)
