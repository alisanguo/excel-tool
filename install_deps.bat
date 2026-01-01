@echo off
REM 安装打包依赖

echo 安装 Excel比对工具 打包依赖...
echo.

pip install -r requirements_build.txt

echo.
echo 安装完成!
echo.
echo 现在可以运行打包脚本:
echo   build.bat
echo.
pause

