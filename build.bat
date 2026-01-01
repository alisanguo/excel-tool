@echo off
REM Excel比对工具 - 打包脚本 (Windows)

echo ======================================
echo Excel比对工具 - 打包脚本
echo ======================================
echo.

REM 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 python
    echo 请先安装 Python 3.7+
    pause
    exit /b 1
)

REM 检查依赖
echo 检查依赖...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo 错误: PyInstaller 未安装
    echo 请运行: pip install pyinstaller
    pause
    exit /b 1
)

python -c "import openpyxl" 2>nul
if errorlevel 1 (
    echo 错误: openpyxl 未安装
    echo 请运行: pip install openpyxl
    pause
    exit /b 1
)

echo √ 依赖检查通过
echo.

REM 运行打包脚本
echo 开始打包...
python build.py

echo.
echo 完成!
pause

