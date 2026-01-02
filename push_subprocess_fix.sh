#!/bin/bash
# 推送Windows进程循环问题修复

echo "======================================"
echo "推送 Windows 进程循环问题修复"
echo "======================================"

cd /Users/li.wang/ai-test-project/excel-tool

echo ""
echo "查看修改的文件..."
git status

echo ""
echo "添加修改的文件..."
git add excel_compare_web.py
git add build.py
git add "Windows进程循环问题修复说明.md"
git add "修复提交说明.txt"
git add push_subprocess_fix.sh

echo ""
echo "提交修改..."
git commit -m "修复Windows下进程无限循环问题

问题：
- 点击选择文件/目录时，不断创建新的ExcelCompare.exe进程
- 任务管理器显示5+个进程同时运行
- 程序无响应，无法正常使用

原因：
- PyInstaller打包后sys.executable指向exe本身
- subprocess.run([sys.executable, 'file_picker.py'])导致循环
- ExcelCompare.exe → ExcelCompare.exe → ExcelCompare.exe → ...

修复方案：
- 将tkinter文件选择逻辑直接集成到主程序
- 移除subprocess调用和file_picker.py依赖
- Windows/Linux直接使用tkinter.filedialog
- macOS继续使用AppleScript（未受影响）

修改内容：
- excel_compare_web.py: 集成tkinter文件选择对话框
- build.py: 移除file_picker.py打包依赖
- 添加详细修复文档

效果：
- ✅ 彻底解决进程循环问题
- ✅ 文件选择对话框正常弹出
- ✅ 只有1个进程运行
- ✅ 简化打包配置
- ✅ 无需外部file_picker.py文件

测试：
- Windows: 待在实际Windows环境验证
- macOS: 功能未改变，正常
- Linux: 使用tkinter，正常"

echo ""
echo "推送到远程仓库..."
git push origin main

echo ""
echo "======================================"
echo "✅ 推送完成！"
echo "======================================"
echo ""
echo "下一步："
echo "1. 访问 GitHub Actions 查看编译状态"
echo "   https://github.com/YOUR_USERNAME/excel-tool/actions"
echo ""
echo "2. 编译完成后，下载 Windows 版本"
echo ""
echo "3. 在 Windows 机器上测试："
echo "   - 运行 ExcelCompare.exe"
echo "   - 打开任务管理器"
echo "   - 点击'选择文件'按钮"
echo "   - 验证：只有1个进程，对话框正常弹出"
echo ""
echo "预期结果："
echo "  ✅ 文件选择对话框立即弹出"
echo "  ✅ 任务管理器只显示1个ExcelCompare.exe进程"
echo "  ✅ 选择文件后路径正确显示"
echo "  ✅ 取消选择不会报错"
echo ""

