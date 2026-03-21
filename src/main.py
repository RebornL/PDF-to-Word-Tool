"""
PDF转Word工具 - 主程序入口
"""

import sys
import os

# 处理PyInstaller打包后的路径
if getattr(sys, 'frozen', False):
    # 打包后的路径
    application_path = sys._MEIPASS
else:
    # 开发环境路径
    application_path = os.path.dirname(os.path.abspath(__file__))

# 添加路径
sys.path.insert(0, application_path)

# 导入并运行主程序
from gui import main

if __name__ == "__main__":
    main()