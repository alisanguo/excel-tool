#!/bin/bash
# 快速提交修复并推送到 GitHub

echo "======================================"
echo "提交 GitHub Actions 修复"
echo "======================================"
echo ""

# 检查 Git 状态
if [ ! -d .git ]; then
    echo "❌ 错误: 当前目录不是 Git 仓库"
    echo ""
    echo "请先初始化 Git 仓库:"
    echo "  git init"
    echo "  git remote add origin https://github.com/你的用户名/仓库名.git"
    exit 1
fi

# 显示变更文件
echo "📁 变更的文件:"
git status --short

echo ""
echo "======================================"
read -p "是否提交这些修复? (y/n): " confirm

if [ "$confirm" != "y" ]; then
    echo "已取消"
    exit 0
fi

# 提交
git add .github/workflows/build.yml
git add build.py
git add GITHUB_ACTIONS_FIX.md
git add 修复说明.txt

git commit -m "Fix: Update GitHub Actions to v4 and improve CI automation

- Update actions/checkout@v3 to @v4
- Update actions/setup-python@v4 to @v5
- Update actions/upload-artifact@v3 to @v4
- Update actions/download-artifact@v3 to @v4
- Update softprops/action-gh-release@v1 to @v2
- Add CI environment detection in build.py
- Fix cross-platform path separators
- Add noconfirm flag for non-interactive builds"

echo ""
echo "✅ 已提交修复"
echo ""

# 推送
read -p "是否推送到 GitHub? (y/n): " push_confirm

if [ "$push_confirm" == "y" ]; then
    echo ""
    echo "正在推送..."
    
    # 获取当前分支名
    branch=$(git branch --show-current)
    
    git push origin "$branch"
    
    if [ $? -eq 0 ]; then
        echo ""
        echo "======================================"
        echo "✅ 推送成功!"
        echo "======================================"
        echo ""
        echo "GitHub Actions 会自动开始编译"
        echo ""
        echo "查看编译进度:"
        echo "  1. 访问你的 GitHub 仓库"
        echo "  2. 点击 Actions 标签"
        echo "  3. 查看运行中的 workflow"
        echo ""
    else
        echo ""
        echo "❌ 推送失败"
        echo "请检查远程仓库地址和权限"
    fi
else
    echo ""
    echo "已取消推送"
    echo ""
    echo "稍后可以手动推送:"
    echo "  git push origin main"
fi

echo ""
echo "完成!"

