#!/bin/bash
# Excel比对工具 - 打包脚本 (macOS/Linux)

echo "======================================"
echo "Excel比对工具 - 打包脚本"
echo "======================================"
echo ""

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 python3"
    exit 1
fi

# 检查依赖
echo "检查依赖..."
python3 -c "import PyInstaller" 2>/dev/null || {
    echo "错误: PyInstaller 未安装"
    echo "请运行: pip install pyinstaller"
    exit 1
}

python3 -c "import openpyxl" 2>/dev/null || {
    echo "错误: openpyxl 未安装"
    echo "请运行: pip install openpyxl"
    exit 1
}

echo "✓ 依赖检查通过"
echo ""

# 运行打包脚本
echo "开始打包..."
python3 build.py

echo ""
echo "完成!"

