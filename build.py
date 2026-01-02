#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel比对工具 - 打包脚本
支持多平台编译成独立可执行文件

使用方法:
    python build.py
    
要求:
    pip install pyinstaller
"""

import os
import sys
import platform
import shutil
import subprocess

# 设置标准输出编码为UTF-8（解决Windows控制台中文输出问题）
if sys.platform == 'win32':
    try:
        # Python 3.7+
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        # Python 3.6及更早版本
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# 设置环境变量，确保子进程也使用UTF-8
os.environ['PYTHONIOENCODING'] = 'utf-8'

# 重写print函数以处理编码错误
import builtins
_original_print = builtins.print

def safe_print(*args, **kwargs):
    """安全打印函数，处理编码错误"""
    try:
        _original_print(*args, **kwargs)
    except UnicodeEncodeError:
        # 如果遇到编码错误，尝试用ASCII安全模式
        safe_args = []
        for arg in args:
            if isinstance(arg, str):
                safe_args.append(arg.encode('ascii', 'replace').decode('ascii'))
            else:
                safe_args.append(arg)
        _original_print(*safe_args, **kwargs)

# 替换内置print函数
builtins.print = safe_print


def get_platform_name():
    """获取平台名称"""
    system = platform.system().lower()
    if system == 'darwin':
        return 'macos'
    elif system == 'windows':
        return 'windows'
    else:
        return 'linux'


def clean_build():
    """清理构建目录"""
    print("清理旧的构建文件...")
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for d in dirs_to_clean:
        if os.path.exists(d):
            shutil.rmtree(d)
            print(f"  删除: {d}")
    
    # 删除 .spec 文件
    for f in os.listdir('.'):
        if f.endswith('.spec'):
            os.remove(f)
            print(f"  删除: {f}")


def check_dependencies():
    """检查依赖"""
    print("\n检查依赖...")
    try:
        import PyInstaller
        print("  ✓ PyInstaller 已安装")
    except ImportError:
        print("  ✗ PyInstaller 未安装")
        print("\n请先安装 PyInstaller:")
        print("  pip install pyinstaller")
        sys.exit(1)
    
    try:
        import openpyxl
        print("  ✓ openpyxl 已安装")
    except ImportError:
        print("  ✗ openpyxl 未安装")
        print("\n请先安装 openpyxl:")
        print("  pip install openpyxl")
        sys.exit(1)


def build_executable():
    """构建可执行文件"""
    platform_name = get_platform_name()
    
    print(f"\n开始构建 ({platform_name})...")
    
    # 确定路径分隔符
    separator = ';' if platform_name == 'windows' else ':'
    
    # PyInstaller 参数
    args = [
        'pyinstaller',
        '--name=ExcelCompare',
        '--onefile',  # 单文件模式
        '--console',  # 显示控制台（方便查看日志）
        '--noconfirm',  # 不询问，直接覆盖
        f'--add-data=file_picker.py{separator}.'  # 包含文件选择器
    ]
    
    # 添加隐藏导入
    hidden_imports = [
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'decimal',
        'tkinter',
        'tkinter.filedialog',
    ]
    
    for imp in hidden_imports:
        args.append(f'--hidden-import={imp}')
    
    # 主程序
    args.append('excel_compare_web.py')
    
    # 执行构建
    print("\n执行 PyInstaller...")
    print(" ".join(args))
    
    result = subprocess.run(args)
    
    if result.returncode == 0:
        print("\n✓ 构建成功!")
        
        # 获取输出文件路径
        if platform_name == 'windows':
            exe_name = 'ExcelCompare.exe'
        else:
            exe_name = 'ExcelCompare'
        
        output_path = os.path.join('dist', exe_name)
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path) / (1024 * 1024)
            print(f"\n可执行文件: {output_path}")
            print(f"文件大小: {file_size:.2f} MB")
            
            # 创建发布目录
            release_dir = f'release_{platform_name}'
            if os.path.exists(release_dir):
                shutil.rmtree(release_dir)
            os.makedirs(release_dir)
            
            # 复制可执行文件
            release_exe = os.path.join(release_dir, exe_name)
            shutil.copy2(output_path, release_exe)
            
            # 复制 README
            if os.path.exists('README.md'):
                shutil.copy2('README.md', os.path.join(release_dir, 'README.md'))
            
            # 创建使用说明
            create_usage_file(release_dir, platform_name)
            
            print(f"\n发布包已创建: {release_dir}/")
            print("\n使用方法:")
            if platform_name == 'windows':
                print(f"  双击运行: {release_dir}\\{exe_name}")
            else:
                print(f"  运行命令: ./{release_dir}/{exe_name}")
        
        return True
    else:
        print("\n✗ 构建失败!")
        return False


def create_usage_file(release_dir, platform_name):
    """创建使用说明文件"""
    usage_text = """Excel比对工具 - 使用说明
====================================

启动方法:
"""
    
    if platform_name == 'windows':
        usage_text += """
Windows:
  1. 双击 ExcelCompare.exe
  2. 浏览器会自动打开 http://localhost:9527
  3. 如果没有自动打开，请手动在浏览器中访问该地址

停止服务:
  关闭命令行窗口即可
"""
    elif platform_name == 'macos':
        usage_text += """
macOS:
  1. 打开终端
  2. 运行: ./ExcelCompare
  3. 浏览器会自动打开 http://localhost:9527
  4. 如果没有自动打开，请手动在浏览器中访问该地址

停止服务:
  在终端按 Ctrl+C

首次运行可能需要授权:
  右键点击 ExcelCompare -> 打开 -> 确认打开
"""
    else:  # linux
        usage_text += """
Linux:
  1. 打开终端
  2. 添加执行权限: chmod +x ExcelCompare
  3. 运行: ./ExcelCompare
  4. 浏览器会自动打开 http://localhost:9527
  5. 如果没有自动打开，请手动在浏览器中访问该地址

停止服务:
  在终端按 Ctrl+C
"""
    
    usage_text += """
功能说明:
  1. 选择工作目录
  2. 上传三个Excel文件（基准文件、输入1、输入2）
  3. 设置颜色阈值（可选）
  4. 点击"开始对比"
  5. 点击"打开结果"查看Excel文件

注意事项:
  - 支持中文文件名和路径
  - 匹配时自动忽略下划线差异
  - 图例显示在Excel右上角
  - 结果文件保存在工作目录下

问题反馈:
  如遇到问题，请查看 README.md 或联系技术支持
"""
    
    usage_file = os.path.join(release_dir, '使用说明.txt')
    with open(usage_file, 'w', encoding='utf-8') as f:
        f.write(usage_text)
    
    print(f"  创建使用说明: {usage_file}")


def main():
    """主函数"""
    print("=" * 60)
    print("Excel比对工具 - 打包脚本")
    print("=" * 60)
    
    # 检查依赖
    check_dependencies()
    
    # 检查是否在 CI 环境中
    is_ci = os.environ.get('CI') == 'true' or os.environ.get('GITHUB_ACTIONS') == 'true'
    
    # 询问是否清理（CI 环境自动清理）
    if is_ci:
        print("\n检测到 CI 环境，自动清理旧的构建文件...")
        clean_build()
    else:
        print("\n是否清理旧的构建文件? (y/n): ", end='')
        if input().lower() == 'y':
            clean_build()
    
    # 构建
    success = build_executable()
    
    if success:
        print("\n" + "=" * 60)
        print("构建完成!")
        print("=" * 60)
    else:
        print("\n构建失败，请检查错误信息")
        sys.exit(1)


if __name__ == '__main__':
    main()

