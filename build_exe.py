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
    
    # 构建配置文件数据参数
    data_args = []
    for config_pattern in config_files:
        config_path = os.path.join(current_dir, config_pattern)
        # 如果是通配符模式
        if '*' in config_pattern:
            matching_files = glob.glob(config_path)
            for file in matching_files:
                if os.path.exists(file):
                    data_args.extend(['--add-data', f'{file}{separator}.'])
        else:
            # 单个文件
            if os.path.exists(config_path):
                data_args.extend(['--add-data', f'{config_path}{separator}.'])
    
    # 定义基本打包参数
    args = [
        'main.py',
        '--noconfirm',
        '--clean',
        '--name=SQLExcelReporter',
        f'--add-data=.env{separator}.',  # 添加.env文件
        '--hidden-import=pandas',
        '--hidden-import=pymysql',
        '--hidden-import=schedule',
        '--hidden-import=python-dotenv',
        '--hidden-import=argparse',  # 添加argparse模块
        '--onefile'
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
