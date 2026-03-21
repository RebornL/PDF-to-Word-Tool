@echo off
chcp 65001 >nul
echo ====================================
echo PDF转Word工具 - 启动程序
echo ====================================
echo.

cd /d "%~dp0"
python src\main.py

pause
