import PyInstaller.__main__
import os
import sys
import argparse

sys.setrecursionlimit(sys.getrecursionlimit() * 5)

def build_executable(target_platform=None):
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 确定目标平台
    if target_platform is None:
        target_platform = 'windows' if sys.platform.startswith('win') else 'linux'
    
    # 选择正确的路径分隔符
    separator = ';' if target_platform == 'windows' else ':'
    
    # 定义配置文件
    config_files = [
        '.env',                    # 默认环境配置
        'config.ini',             # 可选的 INI 配置
        'config.yaml',            # 可选的 YAML 配置
        'config.json',            # 可选的 JSON 配置
        'configs/*.env',          # configs 目录下的所有 .env 文件
        'configs/*.ini',          # configs 目录下的所有 .ini 文件
        'configs/*.yaml',         # configs 目录下的所有 .yaml 文件
        'configs/*.json',         # configs 目录下的所有 .json 文件
    ]
    
    # 定义基本打包参数
    args = [
        'main.py',
        '--noconfirm',
        '--clean',
        '--name=SQLExcelReporter',
        f'--add-data=.env{separator}.',  # 添加.env文件
        '--hidden-import=schedule',
        '--hidden-import=python-dotenv',
        '--hidden-import=argparse',  # 添加argparse模块
        '--hidden-import=oracledb',  # 添加模块
        '--hidden-import=openpyxl',
        '--hidden-import=secrets',
        '--hidden-import=asyncio',
        '--hidden-import=uuid',
        '--exclude-module=PySide6',
        '--exclude-module=PyQt6',
        '--exclude-module=PyQt5',
        '--exclude-module=PySide',
        '--onefile'
    ]

    # 执行打包
    PyInstaller.__main__.run(args)

if __name__ == '__main__':
    build_executable()
