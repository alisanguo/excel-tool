# 跨平台编译指南

在 macOS 上编译 Windows 可执行文件的完整指南。

## 🎯 方案对比

| 方案 | 难度 | 速度 | 推荐度 |
|------|------|------|--------|
| GitHub Actions | ⭐ 简单 | ⭐⭐⭐ 快 | ⭐⭐⭐⭐⭐ 最推荐 |
| Wine | ⭐⭐ 中等 | ⭐⭐ 中 | ⭐⭐⭐ 可用 |
| 虚拟机/双系统 | ⭐⭐⭐ 复杂 | ⭐ 慢 | ⭐⭐ 备选 |
| 远程Windows机器 | ⭐⭐ 中等 | ⭐⭐⭐ 快 | ⭐⭐⭐⭐ 推荐 |

---

## 方案一：GitHub Actions（最推荐）✨

**优点：**
- ✅ 完全免费（公开仓库）
- ✅ 自动化编译
- ✅ 支持所有平台（Windows/macOS/Linux）
- ✅ 无需本地环境配置
- ✅ 可下载编译好的文件

**缺点：**
- ❌ 需要 GitHub 账号
- ❌ 需要上传代码到 GitHub

### 使用步骤

#### 1. 创建 GitHub 仓库

```bash
cd /Users/li.wang/ai-test-project/excel-tool

# 初始化 Git（如果还没有）
git init

# 添加文件
git add .
git commit -m "Initial commit"

# 创建 GitHub 仓库后推送
git remote add origin https://github.com/你的用户名/excel-compare-tool.git
git push -u origin main
```

#### 2. 启用 GitHub Actions

GitHub Actions 配置文件已创建在：`.github/workflows/build.yml`

推送代码后会自动触发编译。

#### 3. 手动触发编译

1. 访问你的 GitHub 仓库
2. 点击 "Actions" 标签
3. 选择 "Build Multi-Platform" 工作流
4. 点击 "Run workflow" 按钮
5. 等待编译完成（约 5-10 分钟）

#### 4. 下载编译结果

1. 在 Actions 页面找到完成的工作流运行
2. 滚动到底部的 "Artifacts" 区域
3. 下载对应平台的文件：
   - `ExcelCompare-windows` (Windows .exe)
   - `ExcelCompare-macos` (macOS)
   - `ExcelCompare-linux` (Linux)

### 自动化发布

创建 Git 标签会自动创建 Release：

```bash
# 创建版本标签
git tag v1.0.0
git push origin v1.0.0

# GitHub 会自动编译并创建 Release
# 访问 Releases 页面下载文件
```

---

## 方案二：使用 Wine（本地编译）🍷

**优点：**
- ✅ 本地编译，无需联网
- ✅ 可重复使用
- ✅ 编译快速

**缺点：**
- ❌ 需要配置 Wine 环境
- ❌ 可能遇到兼容性问题

### 使用步骤

#### 1. 安装 Wine

```bash
# 安装 Homebrew（如果还没有）
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# 安装 Wine
brew install --cask wine-stable

# 验证安装
wine --version
```

#### 2. 安装 Windows 版 Python

```bash
# 下载 Python 3.9 Windows 安装包
curl -O https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe

# 使用 Wine 安装
wine python-3.9.13-amd64.exe

# 安装时注意：
# ✓ 勾选 "Add Python to PATH"
# ✓ 选择 "Install Now"
```

#### 3. 编译 Windows 版本

```bash
cd /Users/li.wang/ai-test-project/excel-tool

# 添加执行权限
chmod +x build_windows_on_mac.sh

# 运行编译脚本
./build_windows_on_mac.sh
```

#### 4. 查看输出

编译完成后，Windows 版本在：`release_windows/ExcelCompare.exe`

### 常见问题

**Q: Wine 安装失败？**
```bash
# 尝试使用 wine-crossover
brew install --cask wine-crossover
```

**Q: Python 安装失败？**
- 确保下载的是 64 位版本
- 尝试使用较旧版本的 Python (如 3.8)

**Q: 编译后的程序无法运行？**
- 在 Windows 机器上测试
- Wine 编译可能存在兼容性问题

---

## 方案三：远程 Windows 机器

**优点：**
- ✅ 编译结果最可靠
- ✅ 速度快
- ✅ 可以测试

**缺点：**
- ❌ 需要 Windows 机器访问权限
- ❌ 需要配置环境

### 使用 AWS/Azure 临时虚拟机

```bash
# 使用 Azure CLI 创建临时 Windows VM
az vm create \
  --resource-group myResourceGroup \
  --name myWinVM \
  --image Win2019Datacenter \
  --admin-username azureuser

# SSH 连接后在 Windows 上执行：
# 1. 安装 Python
# 2. 安装依赖
# 3. 运行 build.bat
# 4. 下载生成的 exe
```

### 使用 GitHub Codespaces

1. 在 GitHub 仓库中创建 Codespace
2. 选择 Windows 环境
3. 运行 `python build.py`

---

## 方案四：虚拟机（备选）

### Parallels Desktop (macOS)

```bash
# 1. 安装 Parallels Desktop
# 2. 创建 Windows 11 虚拟机
# 3. 在虚拟机中：
#    - 安装 Python
#    - 复制项目文件
#    - 运行 build.bat
# 4. 从虚拟机复制出 exe 文件
```

### VirtualBox (免费)

```bash
# 1. 安装 VirtualBox
brew install --cask virtualbox

# 2. 下载 Windows ISO
# 3. 创建虚拟机
# 4. 按上述步骤编译
```

---

## 💡 推荐流程

### 开发阶段
```bash
# 在 macOS 上开发和测试
python excel_compare_web.py
```

### 发布阶段

**方法 A：使用 GitHub Actions（推荐）**
```bash
# 1. 推送代码到 GitHub
git add .
git commit -m "Release v1.0.0"
git tag v1.0.0
git push origin main --tags

# 2. 等待自动编译
# 3. 从 GitHub Releases 下载所有平台版本
```

**方法 B：本地 Wine 编译（快速）**
```bash
# 编译 Windows 版本
./build_windows_on_mac.sh

# 编译 macOS 版本
./build.sh
```

---

## 🔍 验证编译结果

### 在 Windows 上测试

```powershell
# 1. 复制 ExcelCompare.exe 到 Windows 机器
# 2. 双击运行
# 3. 检查：
#    - 是否正常启动
#    - 浏览器是否自动打开
#    - 功能是否正常
#    - 中文是否正常显示
```

### 检查文件信息

```bash
# macOS/Linux
file release_windows/ExcelCompare.exe
# 应该显示: PE32+ executable (console) x86-64, for MS Windows

# 检查文件大小（应该在 20-30 MB）
ls -lh release_windows/ExcelCompare.exe
```

---

## 📋 编译参数说明

```python
# build.py 中的关键参数
'--onefile',           # 单文件模式（推荐）
'--windowed',          # Windows 无控制台窗口
'--add-data',          # 包含额外文件
'--hidden-import',     # 显式导入模块
'--icon',              # 自定义图标（可选）
```

### 优化文件大小

```python
# 在 build.py 中添加
'--exclude-module=matplotlib',  # 排除不需要的模块
'--strip',                      # 去除调试符号
'--upx-dir=/path/to/upx',      # 使用 UPX 压缩
```

---

## 🚨 常见问题

### 1. 编译后文件太大

**原因：** 包含了完整的 Python 解释器和所有依赖

**解决：**
```bash
# 使用 --onedir 模式（分离依赖）
# 使用 UPX 压缩
# 排除不必要的模块
```

### 2. Wine 编译的程序无法运行

**原因：** Wine 不是完美的 Windows 模拟

**解决：**
- 使用 GitHub Actions 编译
- 使用真实 Windows 环境

### 3. 缺少 DLL 文件

**原因：** 某些依赖未正确打包

**解决：**
```python
# 添加 --hidden-import
'--hidden-import=_tkinter',
```

---

## 📚 参考资源

- [PyInstaller 文档](https://pyinstaller.org/)
- [Wine 官网](https://www.winehq.org/)
- [GitHub Actions 文档](https://docs.github.com/en/actions)
- [Homebrew 官网](https://brew.sh/)

---

## 📝 总结

**推荐方案：**

1. **首选：** GitHub Actions - 自动化、可靠、支持所有平台
2. **备选：** Wine - 本地快速编译
3. **最后：** 虚拟机/远程机器 - 最可靠但麻烦

**快速开始：**
```bash
# 推送到 GitHub，自动编译
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/你的用户名/仓库名.git
git push -u origin main

# 或使用 Wine 本地编译
./build_windows_on_mac.sh
```

