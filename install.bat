@echo off
chcp 65001 >nul
echo ====================================
echo PDF转Word工具 - 安装依赖
echo ====================================
echo.

echo 正在安装依赖...
pip install -r requirements.txt

echo.
echo ====================================
echo 安装完成！
echo ====================================
echo.
echo 运行程序请执行: run.bat
echo.
pause
