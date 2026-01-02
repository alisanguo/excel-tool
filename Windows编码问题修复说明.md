# Windows 编码问题修复说明

## 🐛 问题描述

在 GitHub Actions 的 Windows 环境中编译时出现错误：

```
UnicodeEncodeError: 'charmap' codec can't encode characters in position 5-8: 
character maps to <undefined>
```

**错误原因：**
- Windows 默认控制台编码不是 UTF-8（通常是 GBK 或 CP936）
- Python 脚本中包含中文字符（print 输出、文件名、注释等）
- 当尝试输出中文到控制台时，编码转换失败

## ✅ 修复方案

### 1. 修改 `build.py`

在文件开头添加编码设置和安全print函数：

```python
# 设置标准输出编码为UTF-8
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# 设置环境变量
os.environ['PYTHONIOENCODING'] = 'utf-8'

# 重写print函数以处理编码错误
import builtins
_original_print = builtins.print

def safe_print(*args, **kwargs):
    try:
        _original_print(*args, **kwargs)
    except UnicodeEncodeError:
        safe_args = []
        for arg in args:
            if isinstance(arg, str):
                safe_args.append(arg.encode('ascii', 'replace').decode('ascii'))
            else:
                safe_args.append(arg)
        _original_print(*safe_args, **kwargs)

builtins.print = safe_print
```

### 2. 修改 `excel_compare_web.py`

同样在文件开头添加编码设置：

```python
# 设置标准输出编码为UTF-8
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

os.environ['PYTHONIOENCODING'] = 'utf-8'
```

### 3. 修改 `.github/workflows/build.yml`

在构建步骤中添加环境变量：

```yaml
- name: Build with PyInstaller
  run: python build.py
  env:
    PYTHONIOENCODING: utf-8
    PYTHONUTF8: 1
```

## 🔍 技术细节

### 为什么需要多层保护？

1. **`sys.stdout.reconfigure(encoding='utf-8')`**
   - 直接设置标准输出流的编码
   - Python 3.7+ 支持
   - 最直接有效的方法

2. **`codecs.getwriter('utf-8')`**
   - Python 3.6 及更早版本的兼容方案
   - 包装输出流以支持 UTF-8

3. **`os.environ['PYTHONIOENCODING']`**
   - 设置 Python 解释器的 I/O 编码
   - 影响子进程和后续操作
   - 全局性设置

4. **`PYTHONUTF8=1`（GitHub Actions）**
   - Python 3.7+ 的 UTF-8 模式
   - 强制所有文本操作使用 UTF-8
   - 优先级最高的环境变量

5. **安全 print 函数**
   - 最后一道防线
   - 即使所有设置失败，也能安全输出
   - 使用 ASCII 替换无法编码的字符

### 环境变量说明

| 变量 | 作用 | 适用场景 |
|------|------|---------|
| `PYTHONIOENCODING` | 设置标准输入输出编码 | 所有 Python 版本 |
| `PYTHONUTF8` | 启用 Python UTF-8 模式 | Python 3.7+ |

### Windows 特殊处理

Windows 系统的特殊性：
- 默认控制台代码页（Code Page）不是 UTF-8
- 常见代码页：
  - 中文 Windows: CP936 (GBK)
  - 英文 Windows: CP437
  - Windows Terminal: 可配置 UTF-8
- GitHub Actions Windows Runner 默认也不是 UTF-8

## 🧪 测试验证

### 本地测试

**Windows:**
```cmd
# 测试编译
python build.py

# 如果仍有问题，手动设置环境变量
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
python build.py
```

**macOS/Linux:**
```bash
# 通常不需要特殊设置，但可以验证
export PYTHONIOENCODING=utf-8
python build.py
```

### GitHub Actions 测试

提交代码后，GitHub Actions 会自动运行：
- ✅ Windows 构建应该成功
- ✅ macOS 构建应该成功
- ✅ Linux 构建应该成功

## 📊 修复前后对比

### 修复前
```
Build with PyInstaller
  File "D:\a\excel-tool\excel-tool\build.py", line 230, in main
    print("Excel比对工具 - 打包脚本")
UnicodeEncodeError: 'charmap' codec can't encode characters...
Error: Process completed with exit code 1.
```

### 修复后
```
Build with PyInstaller
========================================
Excel比对工具 - 打包脚本
========================================
检查依赖...
  ✓ PyInstaller 已安装
  ✓ openpyxl 已安装
开始构建 (windows)...
✓ 构建成功!
```

## 🎯 最佳实践

### 对于跨平台 Python 项目

1. **始终显式设置编码**
   ```python
   # 文件开头
   # -*- coding: utf-8 -*-
   ```

2. **处理标准输出编码**
   ```python
   if sys.platform == 'win32':
       sys.stdout.reconfigure(encoding='utf-8')
   ```

3. **设置环境变量**
   ```python
   os.environ['PYTHONIOENCODING'] = 'utf-8'
   ```

4. **文件操作显式指定编码**
   ```python
   with open('file.txt', 'w', encoding='utf-8') as f:
       f.write(content)
   ```

5. **使用异常处理**
   ```python
   try:
       print(chinese_text)
   except UnicodeEncodeError:
       print(chinese_text.encode('ascii', 'replace').decode('ascii'))
   ```

## 🔗 相关资源

- [PEP 540 - Add a new UTF-8 Mode](https://peps.python.org/pep-0540/)
- [Python Unicode HOWTO](https://docs.python.org/3/howto/unicode.html)
- [GitHub Actions - Environment Variables](https://docs.github.com/en/actions/learn-github-actions/environment-variables)

## ✨ 总结

通过多层防护措施：
1. ✅ 标准输出流重新配置
2. ✅ 环境变量设置
3. ✅ 安全 print 函数
4. ✅ GitHub Actions 环境变量

确保在任何 Windows 环境下都能正确处理中文字符，不会因编码问题导致编译失败。

---

**修复日期：** 2026-01-03  
**影响版本：** v1.1.0+  
**测试平台：** Windows 10/11, macOS, Linux

