#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文件/目录选择工具 - 独立运行，避免阻塞主服务器
"""

import sys
import os

def pick_file(initial_dir=""):
    """选择文件"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        root = tk.Tk()
        root.withdraw()
        
        # Windows兼容性设置
        if sys.platform == 'win32':
            root.wm_attributes('-topmost', 1)
        else:
            root.attributes('-topmost', True)
        
        root.focus_force()
        root.update()
        
        # 确保初始目录存在
        if initial_dir and not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")
        
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            initialdir=initial_dir or os.path.expanduser("~"),
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        root.destroy()
        
        # 输出结果
        if file_path:
            print(file_path)
        else:
            print("")  # 空字符串表示用户取消
            
    except Exception as e:
        print("ERROR:" + str(e), file=sys.stderr)
        sys.exit(1)

def pick_dir(initial_dir=""):
    """选择目录"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        root = tk.Tk()
        root.withdraw()
        
        # Windows兼容性设置
        if sys.platform == 'win32':
            root.wm_attributes('-topmost', 1)
        else:
            root.attributes('-topmost', True)
        
        root.focus_force()
        root.update()
        
        # 确保初始目录存在
        if initial_dir and not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")
        
        dir_path = filedialog.askdirectory(
            title="选择工作目录",
            initialdir=initial_dir or os.path.expanduser("~")
        )
        
        root.destroy()
        
        # 输出结果
        if dir_path:
            print(dir_path)
        else:
            print("")  # 空字符串表示用户取消
            
    except Exception as e:
        print("ERROR:" + str(e), file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("ERROR:需要指定模式参数", file=sys.stderr)
        sys.exit(1)
    
    mode = sys.argv[1]
    initial = sys.argv[2] if len(sys.argv) > 2 else ""
    
    if mode == 'file':
        pick_file(initial)
    elif mode == 'dir':
        pick_dir(initial)
    else:
        print("ERROR:未知模式 " + mode, file=sys.stderr)
        sys.exit(1)
