#!/bin/bash
# 推送表头显示文件名功能

echo "======================================"
echo "推送 表头显示文件名功能"
echo "======================================"

cd /Users/li.wang/ai-test-project/excel-tool

echo ""
echo "查看修改的文件..."
git status

echo ""
echo "添加修改的文件..."
git add excel_compare_web.py
git add "表头显示文件名-功能说明.md"
git add "更新说明-表头文件名.txt"
git add "测试验证-表头文件名.md"
git add push_header_filename.sh

echo ""
echo "提交修改..."
git commit -m "新增功能：结果文件表头显示实际文件名

功能：
- 表头显示上传文件的实际名称，而不是默认的'A'和'B'
- 自动从文件路径提取文件名（去除扩展名）
- 差额列和图例也使用实际文件名

示例：
- 上传：2024年数据.xlsx 和 2023年数据.xlsx
- 表头：| 指标名称 | 2024年数据 | 2023年数据 | 差额(2024年数据-2023年数据) | 差异% |

技术实现：
- run_compare(): 使用os.path.basename()提取文件名
- _create_result(): 接收文件名参数并在表头中使用
- 支持中文、英文、特殊字符文件名
- 自动移除.xlsx和.xls扩展名

优点：
- ✅ 表头更清晰易懂
- ✅ 一眼看出比对的文件
- ✅ 支持各种文件名格式
- ✅ 向后兼容（默认值为A/B）

文档：
- 详细功能说明.md
- 测试验证指南.md
- 快速更新说明.txt"

echo ""
echo "推送到远程仓库..."
git push origin main

echo ""
echo "======================================"
echo "✅ 推送完成！"
echo "======================================"
echo ""
echo "测试验证："
echo "1. 启动程序: python3 excel_compare_web.py"
echo "2. 访问: http://localhost:9527"
echo "3. 生成测试文件"
echo "4. 上传并比对"
echo "5. 打开结果文件"
echo "6. 验证：表头显示 'test_data_a' 和 'test_data_b'"
echo ""
echo "详细测试步骤请查看："
echo "  测试验证-表头文件名.md"
echo ""

