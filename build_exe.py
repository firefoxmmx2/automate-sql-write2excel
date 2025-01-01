import PyInstaller.__main__
import os
import sys
import argparse

def build_executable(target_platform=None):
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 确定目标平台
    if target_platform is None:
        target_platform = 'windows' if sys.platform.startswith('win') else 'linux'
    
    # 选择正确的路径分隔符
    separator = ';' if target_platform == 'windows' else ':'
    
    # 定义基本打包参数
    args = [
        'main.py',
        '--onefile',
        '--name', f'SQLExcelReporter{"" if target_platform == "linux" else ".exe"}',
        '--add-data', f'{os.path.join(current_dir, ".env")}{separator}.',
        '--collect-all', 'pandas',
        '--collect-all', 'openpyxl',
        '--collect-all', 'pymysql',
        '--collect-all', 'oracledb',
        '--collect-all', 'dotenv',
        '--collect-all', 'schedule',
        '--clean',
    ]

    # 添加平台特定参数
    if target_platform == 'windows':
        args.extend([
            '--target-platform', 'win32',
            '--target-arch', 'x86_64',
        ])
    
    # 执行打包
    PyInstaller.__main__.run(args)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Build executable for different platforms')
    parser.add_argument('--platform', choices=['windows', 'linux'], help='Target platform (windows or linux)')
    args = parser.parse_args()
    
    build_executable(args.platform)
