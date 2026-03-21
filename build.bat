@echo off
chcp 65001 >nul
echo ====================================
echo PDF转Word工具 - 打包为可执行文件
echo ====================================
echo.

echo 正在检查PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
)

echo.
echo 正在打包程序...
pyinstaller --noconfirm --onefile --windowed ^
    --name "PDF转Word工具" ^
    --add-data "src;src" ^
    --hidden-import=pdf2docx ^
    --hidden-import=pdfplumber ^
    --hidden-import=docx ^
    --hidden-import=PyQt5 ^
    --collect-all pdf2docx ^
    --collect-all pdfplumber ^
    src\main.py

echo.
if exist "dist\PDF转Word工具.exe" (
    echo ====================================
    echo 打包完成！
    echo ====================================
    echo.
    echo 可执行文件位置: dist\PDF转Word工具.exe
    echo.
) else (
    echo 打包失败，请检查错误信息。
)

pause
