# GitHub Actions 编译修复说明

## 🐛 问题描述

GitHub Actions 编译失败，错误信息：
```
This request has been automatically failed because it uses a deprecated version of `actions/upload-artifact: v3`
```

## ✅ 已修复的问题

### 1. 更新 Actions 版本
- ✅ `actions/checkout@v3` → `@v4`
- ✅ `actions/setup-python@v4` → `@v5`
- ✅ `actions/upload-artifact@v3` → `@v4`
- ✅ `actions/download-artifact@v3` → `@v4`
- ✅ `softprops/action-gh-release@v1` → `@v2`

### 2. 改进 build.py
- ✅ 添加 CI 环境检测，自动跳过交互式提示
- ✅ 添加 `--noconfirm` 参数，避免构建时询问
- ✅ 修复 Windows 平台的 `--add-data` 路径分隔符
- ✅ 统一使用 `--console` 模式，方便查看日志

### 3. 优化 workflow 配置
- ✅ 修复 Release 文件路径通配符
- ✅ 添加 `fail_on_unmatched_files: false` 避免路径匹配失败

## 🚀 如何使用

### 步骤 1：提交更新后的代码

```bash
cd /Users/li.wang/ai-test-project/excel-tool

git add .
git commit -m "Fix GitHub Actions build issues"
git push origin main
```

### 步骤 2：触发构建

**方法 A：自动触发**
- 推送到 `main` 或 `master` 分支会自动触发

**方法 B：手动触发**
1. 访问 GitHub 仓库
2. 点击 **Actions** 标签
3. 选择 **Build Multi-Platform**
4. 点击 **Run workflow**
5. 点击绿色 **Run workflow** 按钮

### 步骤 3：等待编译完成

- 编译时间：约 5-10 分钟
- 三个平台会并行编译（Windows、macOS、Linux）

### 步骤 4：下载编译结果

1. 在 Actions 页面找到完成的运行
2. 滚动到底部 **Artifacts** 区域
3. 下载你需要的平台：
   - **ExcelCompare-windows** - Windows 可执行文件
   - **ExcelCompare-macos** - macOS 可执行文件
   - **ExcelCompare-linux** - Linux 可执行文件

## 📦 编译产物说明

每个平台的压缩包包含：
```
release_windows/
├── ExcelCompare.exe      # 可执行文件
├── README.md             # 使用说明
└── 使用说明.txt          # 中文说明

release_macos/
├── ExcelCompare          # 可执行文件
├── README.md
└── 使用说明.txt

release_linux/
├── ExcelCompare          # 可执行文件
├── README.md
└── 使用说明.txt
```

## 🏷️ 创建正式版本发布

如果要创建正式的 Release（可以在 Releases 页面看到）：

```bash
# 创建版本标签
git tag v1.0.0
git push origin v1.0.0

# GitHub Actions 会自动：
# 1. 编译三个平台的版本
# 2. 创建 Release
# 3. 上传所有编译产物到 Release 页面
```

访问仓库的 **Releases** 页面即可看到并下载。

## 🔍 验证编译结果

### Windows
```powershell
# 解压后
ExcelCompare.exe

# 应该看到：
# - 控制台窗口打开
# - 显示 "Excel比对工具 - Web界面"
# - 浏览器自动打开 http://localhost:9527
```

### macOS
```bash
# 解压后
chmod +x ExcelCompare
./ExcelCompare

# 首次运行可能需要授权
# 右键 -> 打开 -> 确认打开
```

### Linux
```bash
# 解压后
chmod +x ExcelCompare
./ExcelCompare

# 如果提示缺少依赖：
# sudo apt-get install python3-tk  # Ubuntu/Debian
```

## 📊 编译环境信息

GitHub Actions 使用的环境：

| 平台 | 系统 | Python 版本 |
|------|------|------------|
| Windows | windows-latest (Server 2022) | 3.9 |
| macOS | macos-latest (13.x) | 3.9 |
| Linux | ubuntu-latest (22.04) | 3.9 |

## ❓ 常见问题

### Q: Actions 页面看不到 Artifacts？

A: 检查：
1. 构建是否成功完成（绿色勾）
2. 是否滚动到页面最底部
3. Artifacts 只保留 30 天

### Q: 下载的文件无法运行？

A: 
1. Windows: 右键 -> 属性 -> 解除锁定
2. macOS: 右键 -> 打开（不要双击）
3. Linux: `chmod +x ExcelCompare`

### Q: 想修改编译配置？

A: 编辑 `build.py`：
- 修改程序名称：`--name=你的名称`
- 修改窗口模式：`--console` 或 `--windowed`
- 添加图标：`--icon=icon.ico`
- 排除模块：`--exclude-module=模块名`

### Q: 能否只编译特定平台？

A: 编辑 `.github/workflows/build.yml`：
```yaml
strategy:
  matrix:
    os: [windows-latest]  # 只编译 Windows
```

## 📝 更新日志

### v1.0.1 (修复)
- ✅ 修复 Actions 版本过时问题
- ✅ 改进 CI 环境自动化
- ✅ 优化跨平台编译参数
- ✅ 修复文件路径问题

## 🔗 相关链接

- [GitHub Actions 文档](https://docs.github.com/en/actions)
- [PyInstaller 文档](https://pyinstaller.org/)
- [actions/upload-artifact@v4 变更](https://github.com/actions/upload-artifact/releases/tag/v4.0.0)

## 📞 技术支持

如果还有问题，请检查：
1. Actions 运行日志（点击失败的任务查看详细日志）
2. 确保 `requirements_build.txt` 包含所有依赖
3. 确保 `file_picker.py` 存在于项目根目录

---

**现在可以重新推送代码并触发编译了！** 🎉

