# Excel比对工具 - 打包指南

将Python程序打包成独立的可执行文件，无需Python环境即可运行。

## 📋 准备工作

### 1. 安装依赖

```bash
# 安装打包所需的所有依赖
pip install -r requirements_build.txt
```

或手动安装：

```bash
pip install pyinstaller==5.13.2
pip install openpyxl==3.0.10
```

## 🔨 打包方法

### 方法一：使用打包脚本（推荐）

#### Windows
```bash
# 双击运行或在命令行执行
build.bat
```

#### macOS/Linux
```bash
# 添加执行权限
chmod +x build.sh

# 运行打包脚本
./build.sh
```

### 方法二：使用 Python 脚本

```bash
# 所有平台通用
python build.py
```

## 📦 输出结果

打包完成后会生成以下目录：

```
release_windows/    # Windows 版本
├── ExcelCompare.exe
├── README.md
└── 使用说明.txt

release_macos/      # macOS 版本
├── ExcelCompare
├── README.md
└── 使用说明.txt

release_linux/      # Linux 版本
├── ExcelCompare
├── README.md
└── 使用说明.txt
```

## 🚀 运行打包后的程序

### Windows
1. 双击 `ExcelCompare.exe`
2. 浏览器自动打开 http://localhost:9527

### macOS
```bash
# 首次运行需要授权
# 右键点击 -> 打开 -> 确认打开

# 或在终端运行
./ExcelCompare
```

### Linux
```bash
# 添加执行权限
chmod +x ExcelCompare

# 运行
./ExcelCompare
```

## ⚙️ 打包配置说明

### 文件大小
- Windows: ~20-30 MB
- macOS: ~20-30 MB
- Linux: ~20-30 MB

### 打包模式
- **单文件模式** (`--onefile`): 所有依赖打包成一个exe/可执行文件
- 启动稍慢（需要解压），但分发方便

### 包含的组件
- Python 解释器
- openpyxl 库
- tkinter (文件对话框)
- file_picker.py (辅助脚本)
- HTTP 服务器

## 🔧 自定义打包

### 添加图标

1. 准备图标文件 `icon.ico` (Windows) 或 `icon.icns` (macOS)
2. 修改 `build.py` 中的图标参数：

```python
args = [
    'pyinstaller',
    '--name=ExcelCompare',
    '--onefile',
    '--icon=icon.ico',  # 修改这里
    ...
]
```

### 修改程序名称

修改 `build.py` 中的 `--name` 参数：

```python
'--name=你的程序名',
```

### 多文件模式（启动更快）

将 `--onefile` 改为 `--onedir`：

```python
'--onedir',  # 多文件模式
```

## 📝 常见问题

### Q: 打包后文件太大？
A: 正常现象。包含了Python解释器和所有依赖库。可以考虑：
- 使用 UPX 压缩
- 使用 `--exclude-module` 排除不需要的模块

### Q: 打包后运行报错？
A: 检查：
1. 是否包含了所有必要的文件（file_picker.py）
2. 是否遗漏了隐藏导入（--hidden-import）
3. 在原始Python环境下是否能正常运行

### Q: Windows 杀毒软件报毒？
A: 误报。PyInstaller 打包的程序可能被某些杀毒软件误判。
解决方法：
1. 添加到白名单
2. 使用代码签名证书签名程序

### Q: macOS 提示"无法验证开发者"？
A: 右键点击程序 -> 打开 -> 确认打开
或在终端运行：
```bash
xattr -cr ExcelCompare
```

### Q: Linux 缺少依赖？
A: 某些系统可能缺少 tkinter：
```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# Fedora
sudo dnf install python3-tkinter
```

## 🌐 跨平台编译

注意：只能在对应平台上编译该平台的可执行文件：
- Windows exe 需要在 Windows 上编译
- macOS 可执行文件需要在 macOS 上编译
- Linux 可执行文件需要在 Linux 上编译

### CI/CD 自动化

可以使用 GitHub Actions 等CI/CD工具在不同平台上自动编译：

```yaml
# .github/workflows/build.yml 示例
name: Build
on: [push]
jobs:
  build:
    runs-on: ${{ matrix.os }}
    strategy:
      matrix:
        os: [ubuntu-latest, windows-latest, macos-latest]
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-python@v2
      - run: pip install -r requirements_build.txt
      - run: python build.py
```

## 📚 更多资源

- [PyInstaller 官方文档](https://pyinstaller.org/)
- [Python 打包指南](https://packaging.python.org/)

## 📄 许可证

MIT License

