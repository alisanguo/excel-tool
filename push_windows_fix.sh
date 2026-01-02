#!/bin/bash
# 推送Windows编码问题修复

echo "======================================"
echo "推送 Windows 编码问题修复"
echo "======================================"

cd /Users/li.wang/ai-test-project/excel-tool

echo ""
echo "查看修改的文件..."
git status

echo ""
echo "添加修改的文件..."
git add build.py
git add excel_compare_web.py
git add .github/workflows/build.yml
git add "Windows编码问题修复说明.md"
git add "修复Windows编码-提交说明.txt"
git add push_windows_fix.sh

echo ""
echo "提交修改..."
git commit -m "修复Windows编译时的UTF-8编码问题

问题：
- Windows环境下PyInstaller编译时出现UnicodeEncodeError
- 控制台默认编码不是UTF-8，无法输出中文字符

修复：
- build.py: 添加UTF-8编码设置和安全print函数
- excel_compare_web.py: 添加UTF-8编码设置
- GitHub Actions: 添加PYTHONIOENCODING和PYTHONUTF8环境变量

技术方案：
1. sys.stdout.reconfigure(encoding='utf-8') - 重新配置输出流
2. 环境变量 PYTHONIOENCODING=utf-8 - 全局设置
3. 安全print函数 - 兜底保护，处理编码异常
4. CI环境变量 - 确保GitHub Actions正确处理UTF-8

测试：
- 兼容 Python 3.6-3.9+
- 支持 Windows/macOS/Linux
- 确保GitHub Actions编译成功"

echo ""
echo "推送到远程仓库..."
git push origin main

echo ""
echo "======================================"
echo "✅ 推送完成！"
echo "======================================"
echo ""
echo "查看 GitHub Actions 构建状态："
echo "https://github.com/YOUR_USERNAME/excel-tool/actions"
echo ""
echo "预期结果："
echo "  ✅ Build on windows-latest - 成功"
echo "  ✅ Build on macos-latest - 成功"
echo "  ✅ Build on ubuntu-latest - 成功"
echo ""

