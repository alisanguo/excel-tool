#!/bin/bash
# 在 macOS 上编译 Windows 可执行文件
# 使用 Wine 运行 Windows 版 Python

echo "======================================"
echo "在 macOS 上编译 Windows 版本"
echo "======================================"
echo ""

# 检查是否安装了 Wine
if ! command -v wine &> /dev/null; then
    echo "❌ 未检测到 Wine"
    echo ""
    echo "请先安装 Wine:"
    echo "  brew install --cask wine-stable"
    echo ""
    echo "或使用 Homebrew:"
    echo "  1. 安装 Homebrew: https://brew.sh"
    echo "  2. 运行: brew install --cask wine-stable"
    echo ""
    exit 1
fi

echo "✓ 检测到 Wine"
echo ""

# Wine Python 安装路径
WINE_PYTHON="$HOME/.wine/drive_c/Python39"
WINE_PYTHON_EXE="$WINE_PYTHON/python.exe"

# 检查是否安装了 Windows 版 Python
if [ ! -f "$WINE_PYTHON_EXE" ]; then
    echo "❌ 未检测到 Wine 中的 Python"
    echo ""
    echo "请安装 Windows 版 Python:"
    echo "  1. 下载 Python 3.9 Windows 安装包"
    echo "     https://www.python.org/downloads/windows/"
    echo "  2. 使用 Wine 运行安装包:"
    echo "     wine python-3.9.x.exe"
    echo "  3. 安装时选择 'Add to PATH'"
    echo ""
    exit 1
fi

echo "✓ 检测到 Wine Python"
echo ""

# 安装依赖
echo "安装依赖..."
wine "$WINE_PYTHON_EXE" -m pip install --upgrade pip
wine "$WINE_PYTHON_EXE" -m pip install -r requirements_build.txt

if [ $? -ne 0 ]; then
    echo "❌ 依赖安装失败"
    exit 1
fi

echo "✓ 依赖安装完成"
echo ""

# 编译
echo "开始编译 Windows 版本..."
wine "$WINE_PYTHON_EXE" -m PyInstaller \
    --name=ExcelCompare \
    --onefile \
    --windowed \
    --add-data="file_picker.py;." \
    --hidden-import=openpyxl \
    --hidden-import=openpyxl.styles \
    --hidden-import=openpyxl.utils \
    --hidden-import=decimal \
    --hidden-import=tkinter \
    --hidden-import=tkinter.filedialog \
    excel_compare_web.py

if [ $? -eq 0 ]; then
    echo ""
    echo "✓ 编译成功!"
    echo ""
    
    # 创建发布目录
    RELEASE_DIR="release_windows"
    mkdir -p "$RELEASE_DIR"
    
    # 复制文件
    if [ -f "dist/ExcelCompare.exe" ]; then
        cp dist/ExcelCompare.exe "$RELEASE_DIR/"
        cp README.md "$RELEASE_DIR/" 2>/dev/null || true
        
        # 创建使用说明
        cat > "$RELEASE_DIR/使用说明.txt" << 'EOF'
Excel比对工具 - 使用说明
====================================

启动方法:
  双击 ExcelCompare.exe

浏览器会自动打开 http://localhost:9527
如果没有自动打开，请手动在浏览器中访问该地址

停止服务:
  关闭命令行窗口即可

功能说明:
  1. 选择工作目录
  2. 上传三个Excel文件
  3. 点击"开始对比"
  4. 查看结果

支持:
  - 中文文件名和路径
  - 自动忽略下划线差异
  - 图例显示在右上角
EOF
        
        echo "发布包已创建: $RELEASE_DIR/"
        echo ""
        ls -lh "$RELEASE_DIR/ExcelCompare.exe"
    else
        echo "❌ 未找到编译输出"
        exit 1
    fi
else
    echo ""
    echo "❌ 编译失败"
    exit 1
fi

echo ""
echo "======================================"
echo "完成!"
echo "======================================"

