# 在 Mac 上编译 Windows 版本 - 快速指南

## 🎯 推荐方案：GitHub Actions（最简单）

### 步骤 1：上传到 GitHub

```bash
cd /Users/li.wang/ai-test-project/excel-tool

# 初始化 Git
git init
git add .
git commit -m "Initial commit"

# 创建 GitHub 仓库后（在 GitHub 网站上创建）
git remote add origin https://github.com/你的用户名/仓库名.git
git push -u origin main
```

### 步骤 2：触发自动编译

1. 访问 GitHub 仓库页面
2. 点击 **Actions** 标签
3. 点击左侧 **Build Multi-Platform**
4. 点击右上角 **Run workflow** → **Run workflow**
5. 等待 5-10 分钟

### 步骤 3：下载编译结果

1. 编译完成后，在 Actions 页面找到完成的运行
2. 滚动到底部 **Artifacts** 区域
3. 下载 **ExcelCompare-windows** 压缩包
4. 解压后得到 `ExcelCompare.exe`

✅ **完成！** 无需任何本地配置

---

## 🍷 备选方案：使用 Wine（本地编译）

### 前置要求

需要先安装 Wine 和 Windows 版 Python（只需配置一次）

### 一键编译

```bash
cd /Users/li.wang/ai-test-project/excel-tool

# 运行脚本（会自动检查环境）
./build_windows_on_mac.sh
```

输出文件：`release_windows/ExcelCompare.exe`

---

## 📋 方案对比

| 方案 | 优点 | 缺点 | 推荐度 |
|------|------|------|--------|
| **GitHub Actions** | 免费、自动化、无需配置 | 需要 GitHub 账号 | ⭐⭐⭐⭐⭐ |
| **Wine** | 本地快速、可离线 | 需要配置环境 | ⭐⭐⭐ |

---

## 🔧 Wine 环境配置（首次使用）

### 1. 安装 Wine

```bash
# 安装 Homebrew（如果没有）
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# 安装 Wine
brew install --cask wine-stable

# 验证（等待初始化完成）
wine --version
```

### 2. 下载并安装 Windows 版 Python

```bash
# 下载 Python 3.9 Windows 安装包
curl -O https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe

# 使用 Wine 安装
wine python-3.9.13-amd64.exe
```

**安装时注意：**
- ✅ 勾选 "Add Python to PATH"
- ✅ 选择 "Install Now"
- 等待安装完成（可能需要几分钟）

### 3. 编译

```bash
cd /Users/li.wang/ai-test-project/excel-tool
./build_windows_on_mac.sh
```

---

## ❓ 常见问题

### Q: 我不想用 GitHub，有其他办法吗？

A: 可以使用 Wine（见上方），或者：
- 找一台 Windows 电脑/虚拟机
- 使用云端 Windows 服务器

### Q: GitHub Actions 免费吗？

A: 公开仓库完全免费，私有仓库每月有免费额度

### Q: Wine 编译的程序可靠吗？

A: 大部分情况可用，但建议在真实 Windows 上测试

### Q: 可以同时编译 Windows/Mac/Linux 版本吗？

A: 可以！GitHub Actions 会自动编译所有平台

---

## 📖 详细文档

查看完整指南：`CROSS_PLATFORM_BUILD.md`

---

## 🚀 最快方式（推荐）

```bash
# 1. 上传到 GitHub
git init
git add .
git commit -m "Initial"
# 在 GitHub 创建仓库后
git remote add origin https://github.com/你的用户名/仓库名.git
git push -u origin main

# 2. GitHub 网站操作
# Actions → Build Multi-Platform → Run workflow

# 3. 下载编译好的文件
# Actions → 完成的运行 → Artifacts → 下载
```

**总耗时：约 10 分钟（大部分时间在自动编译）**

