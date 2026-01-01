#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel比对工具 - Web界面版
Python 3.7.1 兼容
启动后在浏览器中打开 http://localhost:8080
"""

import os
import sys
import json
import webbrowser
import threading
import subprocess
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import parse_qs, unquote
from decimal import Decimal, ROUND_HALF_UP

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# 全局配置
WORK_DIR = os.getcwd()
PORT = 9527

HTML_TEMPLATE = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Excel比对工具</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh; padding: 30px;
        }
        .container { 
            max-width: 700px; margin: 0 auto; 
            background: white; border-radius: 16px; 
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 35px; 
        }
        h1 { 
            text-align: center; color: #333; margin-bottom: 30px;
            font-size: 28px; font-weight: 600;
        }
        .section { 
            background: #f8f9fa; border-radius: 12px; 
            padding: 20px; margin-bottom: 20px; 
        }
        .section-title { 
            font-size: 15px; font-weight: 600; color: #555; 
            margin-bottom: 15px; display: flex; align-items: center;
        }
        .section-title::before {
            content: ""; width: 4px; height: 18px; 
            background: #667eea; border-radius: 2px; margin-right: 10px;
        }
        .form-row { 
            display: flex; align-items: center; margin-bottom: 12px; 
        }
        .form-row:last-child { margin-bottom: 0; }
        label { 
            width: 130px; font-size: 14px; color: #444; font-weight: 500;
        }
        input[type="text"], input[type="file"] { 
            flex: 1; padding: 10px 14px; border: 2px solid #e0e0e0; 
            border-radius: 8px; font-size: 14px; transition: border-color 0.2s;
        }
        input[type="text"]:focus { border-color: #667eea; outline: none; }
        input[type="number"] {
            width: 80px; padding: 8px 12px; border: 2px solid #e0e0e0;
            border-radius: 8px; font-size: 14px; text-align: center;
        }
        .color-row { display: flex; align-items: center; margin-bottom: 10px; }
        .color-box { 
            width: 24px; height: 24px; border-radius: 4px; 
            margin-right: 12px; border: 1px solid #ccc;
        }
        .green-box { background: #90EE90; }
        .red-box { background: #FF6B6B; }
        .white-box { background: #fff; }
        .color-text { font-size: 14px; color: #555; }
        .btn-row { 
            display: flex; gap: 12px; margin-top: 25px; flex-wrap: wrap;
        }
        button { 
            padding: 12px 24px; border: none; border-radius: 8px; 
            font-size: 14px; font-weight: 600; cursor: pointer; 
            transition: transform 0.1s, box-shadow 0.2s;
        }
        button:hover { transform: translateY(-1px); }
        button:active { transform: translateY(0); }
        .btn-primary { 
            background: linear-gradient(135deg, #667eea, #764ba2); 
            color: white; box-shadow: 0 4px 15px rgba(102,126,234,0.4);
        }
        .btn-secondary { background: #6c757d; color: white; }
        .btn-success { background: #28a745; color: white; }
        .log-box { 
            background: #1e1e1e; color: #0f0; border-radius: 8px; 
            padding: 15px; font-family: "Courier New", monospace; 
            font-size: 13px; height: 150px; overflow-y: auto;
            white-space: pre-wrap;
        }
        .file-input-wrapper {
            flex: 1; display: flex; gap: 8px;
        }
        .file-path {
            flex: 1; padding: 10px 14px; border: 2px solid #e0e0e0;
            border-radius: 8px; font-size: 13px; background: #fff;
            overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }
        .btn-browse {
            padding: 10px 16px; background: #e9ecef; border: 2px solid #e0e0e0;
            border-radius: 8px; cursor: pointer; font-size: 13px;
        }
        .btn-browse:hover { background: #dee2e6; }
        .hidden-input { display: none; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Excel 数据比对工具</h1>
        
        <div class="section">
            <div class="section-title">工作目录</div>
            <div class="form-row">
                <label>目录路径:</label>
                <div class="file-input-wrapper">
                    <input type="text" id="workDir" value="''' + WORK_DIR.replace('\\', '\\\\').replace("'", "\\'") + '''">
                    <button class="btn-browse" onclick="browseDir()">选择目录</button>
                </div>
            </div>
        </div>
        
        <div class="section">
            <div class="section-title">文件选择</div>
            <div class="form-row">
                <label>上传基准文件:</label>
                <div class="file-input-wrapper">
                    <input type="text" id="baseFile" placeholder="选择基准匹配列文件 (.xlsx)">
                    <button class="btn-browse" onclick="browseFile('baseFile')">选择文件</button>
                </div>
            </div>
            <div class="form-row">
                <label>上传输入1文件:</label>
                <div class="file-input-wrapper">
                    <input type="text" id="dataAFile" placeholder="选择输入1数据文件 (.xlsx)">
                    <button class="btn-browse" onclick="browseFile('dataAFile')">选择文件</button>
                </div>
            </div>
            <div class="form-row">
                <label>上传输入2文件:</label>
                <div class="file-input-wrapper">
                    <input type="text" id="dataBFile" placeholder="选择输入2数据文件 (.xlsx)">
                    <button class="btn-browse" onclick="browseFile('dataBFile')">选择文件</button>
                </div>
            </div>
            <div class="form-row">
                <label>输出文件名:</label>
                <input type="text" id="outputFile" value="compare_result.xlsx">
            </div>
        </div>
        
        <div class="section">
            <div class="section-title">颜色阈值设置 (差异百分比绝对值)</div>
            <div class="color-row">
                <div class="color-box green-box"></div>
                <span class="color-text">绿色: 差异 < </span>
                <input type="number" id="greenTh" value="1.0" step="0.1" style="margin: 0 5px;">
                <span class="color-text">%</span>
            </div>
            <div class="color-row">
                <div class="color-box red-box"></div>
                <span class="color-text">红色: </span>
                <input type="number" id="redMin" value="1.0" step="0.1" style="margin: 0 5px; width: 70px;">
                <span class="color-text">% ≤ 差异 ≤ </span>
                <input type="number" id="redMax" value="100.0" step="1" style="margin: 0 5px; width: 70px;">
                <span class="color-text">%</span>
            </div>
            <div class="color-row">
                <div class="color-box white-box"></div>
                <span class="color-text">无色: 差异 > 红色上限</span>
            </div>
        </div>
        
        <div class="btn-row">
            <button class="btn-secondary" onclick="generateTest()">生成测试文件</button>
            <button class="btn-primary" onclick="runCompare()">🚀 开始对比</button>
            <button class="btn-success" onclick="openResult()">打开结果</button>
            <button class="btn-secondary" onclick="openDir()">打开目录</button>
        </div>
        
        <div class="section" style="margin-top: 20px;">
            <div class="section-title">运行日志</div>
            <div class="log-box" id="logBox">欢迎使用Excel比对工具!
步骤: 1.设置目录 → 2.输入文件路径 → 3.点击开始对比

提示: 请直接输入文件的完整路径，或先点击"生成测试文件"</div>
        </div>
    </div>
    
    <script>
        function log(msg) {
            const box = document.getElementById('logBox');
            box.textContent += '\\n' + msg;
            box.scrollTop = box.scrollHeight;
        }
        
        function clearLog() {
            document.getElementById('logBox').textContent = '';
        }
        
        async function api(action, data) {
            try {
                const resp = await fetch('/api', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({action, ...data})
                });
                return await resp.json();
            } catch(e) {
                return {success: false, message: '请求失败: ' + e.message};
            }
        }
        
        async function generateTest() {
            log('\\n生成测试文件...');
            const workDir = document.getElementById('workDir').value;
            const result = await api('generate_test', {workDir});
            if (result.success) {
                log(result.message);
                document.getElementById('baseFile').value = result.baseFile;
                document.getElementById('dataAFile').value = result.dataAFile;
                document.getElementById('dataBFile').value = result.dataBFile;
                log('文件路径已自动填充!');
            } else {
                log('错误: ' + result.message);
            }
        }
        
        async function runCompare() {
            const data = {
                workDir: document.getElementById('workDir').value,
                baseFile: document.getElementById('baseFile').value,
                dataAFile: document.getElementById('dataAFile').value,
                dataBFile: document.getElementById('dataBFile').value,
                outputFile: document.getElementById('outputFile').value,
                greenTh: parseFloat(document.getElementById('greenTh').value),
                redMin: parseFloat(document.getElementById('redMin').value),
                redMax: parseFloat(document.getElementById('redMax').value)
            };
            
            if (!data.baseFile) { alert('请输入基准文件路径'); return; }
            if (!data.dataAFile) { alert('请输入输入1文件路径'); return; }
            if (!data.dataBFile) { alert('请输入输入2文件路径'); return; }
            
            log('\\n========================================');
            log('开始对比...');
            log('阈值: 绿<' + data.greenTh + '%, 红' + data.redMin + '%-' + data.redMax + '%');
            
            const result = await api('compare', data);
            if (result.success) {
                log(result.message);
                alert('对比完成!');
            } else {
                log('错误: ' + result.message);
                alert('对比失败: ' + result.message);
            }
        }
        
        async function openResult() {
            const workDir = document.getElementById('workDir').value;
            const outputFile = document.getElementById('outputFile').value;
            await api('open_file', {path: workDir + '/' + outputFile});
        }
        
        async function openDir() {
            const workDir = document.getElementById('workDir').value;
            await api('open_dir', {path: workDir});
        }
        
        async function browseFile(inputId) {
            log('正在打开文件选择对话框...');
            const workDir = document.getElementById('workDir').value;
            const result = await api('browse_file', {workDir});
            if (result.success && result.path) {
                document.getElementById(inputId).value = result.path;
                log('已选择: ' + result.path);
            } else if (result.message) {
                log(result.message);
            }
        }
        
        async function browseDir() {
            log('正在打开目录选择对话框...');
            const result = await api('browse_dir', {});
            if (result.success && result.path) {
                document.getElementById('workDir').value = result.path;
                log('工作目录: ' + result.path);
            }
        }
    </script>
</body>
</html>
'''


class RequestHandler(BaseHTTPRequestHandler):
    """HTTP请求处理"""
    
    def log_message(self, format, *args):
        pass  # 禁用默认日志
    
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(HTML_TEMPLATE.encode('utf-8'))
    
    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length).decode('utf-8')
        
        try:
            data = json.loads(body)
            action = data.get('action', '')
            
            if action == 'generate_test':
                result = self.generate_test(data.get('workDir', WORK_DIR))
            elif action == 'compare':
                result = self.run_compare(data)
            elif action == 'open_file':
                result = self.open_file(data.get('path', ''))
            elif action == 'open_dir':
                result = self.open_dir(data.get('path', ''))
            elif action == 'browse_file':
                result = self.browse_file_dialog(data.get('workDir', WORK_DIR))
            elif action == 'browse_dir':
                result = self.browse_dir_dialog()
            else:
                result = {'success': False, 'message': '未知操作'}
                
        except Exception as e:
            result = {'success': False, 'message': str(e)}
        
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(result, ensure_ascii=False).encode('utf-8'))
    
    def generate_test(self, workdir):
        """生成测试文件"""
        if not OPENPYXL_OK:
            return {'success': False, 'message': '缺少openpyxl库'}
        
        if not os.path.exists(workdir):
            return {'success': False, 'message': '工作目录不存在'}
        
        try:
            # 基准文件
            wb = Workbook()
            ws = wb.active
            ws.cell(row=1, column=1, value="指标名称")
            indicators = [
                "正常_完全相同", "正常_小差异_0.5%", "正常_临界_1%", "正常_中等_5%",
                "正常_较大_50%", "正常_超大_150%", "特殊_B为零", "特殊_负数",
                "缺失_A无数据", "缺失_B无数据", "缺失_都无数据"
            ]
            for i, name in enumerate(indicators, 2):
                ws.cell(row=i, column=1, value=name)
            base_path = os.path.join(workdir, "test_base.xlsx")
            wb.save(base_path)
            
            # 数据A
            wb = Workbook()
            ws = wb.active
            data_a = [
                ("正常_完全相同", 1000000), ("正常_小差异_0.5%", 1005000),
                ("正常_临界_1%", 1010000), ("正常_中等_5%", 1050000),
                ("正常_较大_50%", 1500000), ("正常_超大_150%", 2500000),
                ("特殊_B为零", 100), ("特殊_负数", -500),
                ("缺失_A无数据", None), ("缺失_都无数据", None)
            ]
            for col, (h, v) in enumerate(data_a, 1):
                ws.cell(row=1, column=col, value=h)
                ws.cell(row=2, column=col, value=v)
            data_a_path = os.path.join(workdir, "test_data_a.xlsx")
            wb.save(data_a_path)
            
            # 数据B
            wb = Workbook()
            ws = wb.active
            data_b = [
                ("正常_完全相同", 1000000), ("正常_小差异_0.5%", 1000000),
                ("正常_临界_1%", 1000000), ("正常_中等_5%", 1000000),
                ("正常_较大_50%", 1000000), ("正常_超大_150%", 1000000),
                ("特殊_B为零", 0), ("特殊_负数", -400),
                ("缺失_B无数据", None), ("缺失_都无数据", None)
            ]
            for col, (h, v) in enumerate(data_b, 1):
                ws.cell(row=1, column=col, value=h)
                ws.cell(row=2, column=col, value=v)
            data_b_path = os.path.join(workdir, "test_data_b.xlsx")
            wb.save(data_b_path)
            
            return {
                'success': True,
                'message': '测试文件已生成:\n  - test_base.xlsx\n  - test_data_a.xlsx\n  - test_data_b.xlsx',
                'baseFile': base_path,
                'dataAFile': data_a_path,
                'dataBFile': data_b_path
            }
            
        except Exception as e:
            return {'success': False, 'message': str(e)}
    
    def run_compare(self, data):
        """运行对比"""
        if not OPENPYXL_OK:
            return {'success': False, 'message': '缺少openpyxl库'}
        
        try:
            workdir = data.get('workDir', WORK_DIR)
            base_file = data.get('baseFile', '')
            data_a_file = data.get('dataAFile', '')
            data_b_file = data.get('dataBFile', '')
            output_file = data.get('outputFile', 'compare_result.xlsx')
            green_th = float(data.get('greenTh', 1.0))
            red_min = float(data.get('redMin', 1.0))
            red_max = float(data.get('redMax', 100.0))
            
            # 读取基准
            base_names = self._read_base(base_file)
            
            # 读取数据
            data_a = self._read_horizontal(data_a_file)
            data_b = self._read_horizontal(data_b_file)
            
            # 生成结果
            output_path = os.path.join(workdir, output_file)
            self._create_result(output_path, base_names, data_a, data_b, green_th, red_min, red_max)
            
            return {
                'success': True, 
                'message': '基准: {} 个指标\n输入1: {} 个数据\n输入2: {} 个数据\n========================================\n[完成] 结果已保存: {}'.format(
                    len(base_names), len(data_a), len(data_b), output_file
                )
            }
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {'success': False, 'message': str(e)}
    
    def _read_base(self, path):
        # 处理中文路径
        if sys.platform == 'win32' and isinstance(path, str):
            # Windows上确保路径是Unicode字符串
            path = os.path.normpath(path)
        
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        names = []
        for row in range(2, ws.max_row + 1):
            v = ws.cell(row=row, column=1).value
            if v:
                names.append(str(v).strip())
        wb.close()
        return names
    
    def _read_horizontal(self, path):
        # 处理中文路径
        if sys.platform == 'win32' and isinstance(path, str):
            # Windows上确保路径是Unicode字符串
            path = os.path.normpath(path)
        
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        data = {}
        for col in range(1, ws.max_column + 1):
            h = ws.cell(row=1, column=col).value
            if h:
                # 保存原始key和标准化key的映射
                original_key = str(h).strip()
                data[original_key] = ws.cell(row=2, column=col).value
        wb.close()
        return data
    
    def _normalize_key(self, key):
        """标准化指标名称，只忽略下划线"""
        return key.replace('_', '').lower()
    
    def _find_value(self, data_dict, target_key):
        """根据标准化规则查找值，忽略下划线差异"""
        # 先尝试精确匹配
        if target_key in data_dict:
            return data_dict[target_key]
        
        # 标准化后模糊匹配（只忽略下划线）
        normalized_target = self._normalize_key(target_key)
        for key, value in data_dict.items():
            if self._normalize_key(key) == normalized_target:
                return value
        
        return None
    
    def _parse_num(self, v):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return Decimal(str(v))
        s = str(v).strip().replace(',', '').replace(' ', '')
        if not s or s.lower() in ['error', '#value!', 'none', 'null']:
            return None
        try:
            return Decimal(s)
        except:
            return None
    
    def _create_result(self, output, names, data_a, data_b, green_th, red_min, red_max):
        GREEN = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        RED = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
        HEADER = PatternFill(start_color="DCDCDC", end_color="DCDCDC", fill_type="solid")
        LEGEND_FILL = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        wb = Workbook()
        ws = wb.active
        ws.title = "比对结果"
        
        # 表头（第1行）
        for col, h in enumerate(["指标名称", "A", "B", "差额", "差异%"], 1):
            c = ws.cell(row=1, column=col, value=h)
            c.fill = HEADER
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal='center')
            c.border = border
        
        # 图例放在右上角 G1:H2（与表头同行及下一行）
        legend_col = 7  # G列
        cell_g1 = ws.cell(row=1, column=legend_col, value="差异 >= {}%".format(red_min))
        cell_g1.border = border
        cell_g1.fill = LEGEND_FILL
        cell_g1.alignment = Alignment(horizontal='left')
        cell_g1.font = Font(size=10)
        
        cell_h1 = ws.cell(row=1, column=legend_col+1, value="红色")
        cell_h1.fill = RED
        cell_h1.border = border
        cell_h1.alignment = Alignment(horizontal='center')
        cell_h1.font = Font(size=10)
        
        cell_g2 = ws.cell(row=2, column=legend_col, value="差异 < {}%".format(green_th))
        cell_g2.border = border
        cell_g2.fill = LEGEND_FILL
        cell_g2.alignment = Alignment(horizontal='left')
        cell_g2.font = Font(size=10)
        
        cell_h2 = ws.cell(row=2, column=legend_col+1, value="绿色")
        cell_h2.fill = GREEN
        cell_h2.border = border
        cell_h2.alignment = Alignment(horizontal='center')
        cell_h2.font = Font(size=10)
        
        # 数据行（从第2行开始）
        current_row = 2
        for name in names:
            ws.cell(row=current_row, column=1, value=name).border = border
            
            # 使用模糊匹配查找值
            va = self._find_value(data_a, name)
            vb = self._find_value(data_b, name)
            pa = self._parse_num(va)
            pb = self._parse_num(vb)
            
            if pa is not None:
                ws.cell(row=current_row, column=2, value=float(pa)).border = border
            else:
                ws.cell(row=current_row, column=2, value="error").border = border
                
            if pb is not None:
                ws.cell(row=current_row, column=3, value=float(pb)).border = border
            else:
                ws.cell(row=current_row, column=3, value="error").border = border
                
            if pa is not None and pb is not None:
                diff = pa - pb
                ws.cell(row=current_row, column=4, value=float(diff.quantize(Decimal('0.0001')))).border = border
            else:
                ws.cell(row=current_row, column=4, value="#VALUE!").border = border
                diff = None
                
            cell = ws.cell(row=current_row, column=5)
            cell.border = border
            if diff is not None and pb is not None and pb != 0:
                pct = (diff / pb) * 100
                pct_val = float(pct.quantize(Decimal('0.0001')))
                cell.value = "{}%".format(pct_val)
                
                abs_pct = abs(pct)
                if abs_pct < green_th:
                    cell.fill = GREEN
                elif red_min <= abs_pct <= red_max:
                    cell.fill = RED
            else:
                cell.value = "#VALUE!"
            
            current_row += 1
                
        # 调整列宽
        for col, w in enumerate([22, 18, 18, 16, 16, 16, 10], 1):
            ws.column_dimensions[get_column_letter(col)].width = w
        
        # 保存文件，处理中文路径编码
        try:
            wb.save(output)
        except Exception as e:
            # 如果保存失败，尝试用不同的编码
            if sys.platform == 'win32':
                # Windows上尝试使用UTF-8
                output_bytes = output.encode('utf-8')
                wb.save(output_bytes.decode('utf-8'))
            else:
                raise e
    
    def open_file(self, path):
        if os.path.exists(path):
            if sys.platform == 'darwin':
                os.system('open "{}"'.format(path))
            elif sys.platform == 'win32':
                os.system('start "" "{}"'.format(path))
            else:
                os.system('xdg-open "{}"'.format(path))
            return {'success': True}
        return {'success': False, 'message': '文件不存在'}
    
    def open_dir(self, path):
        if os.path.exists(path):
            if sys.platform == 'darwin':
                os.system('open "{}"'.format(path))
            elif sys.platform == 'win32':
                os.system('explorer "{}"'.format(path))
            else:
                os.system('xdg-open "{}"'.format(path))
            return {'success': True}
        return {'success': False, 'message': '目录不存在'}
    
    def browse_file_dialog(self, initial_dir):
        """打开文件选择对话框"""
        try:
            if sys.platform == 'darwin':
                # macOS: 使用osascript
                script = '''
                tell application "System Events"
                    activate
                    set theFile to choose file with prompt "选择Excel文件" of type {"xlsx", "xls"}
                    return POSIX path of theFile
                end tell
                '''
                result = subprocess.run(['osascript', '-e', script], 
                                        capture_output=True, text=True, timeout=60)
                if result.returncode == 0 and result.stdout.strip():
                    return {'success': True, 'path': result.stdout.strip()}
                return {'success': False, 'message': '未选择文件'}
            else:
                # Windows/Linux: 使用独立进程运行tkinter
                script_dir = os.path.dirname(os.path.abspath(__file__))
                picker_script = os.path.join(script_dir, 'file_picker.py')
                
                # 确保初始目录存在
                if not initial_dir or not os.path.exists(initial_dir):
                    initial_dir = os.getcwd()
                
                # Windows上隐藏控制台窗口
                kwargs = {
                    'capture_output': True,
                    'text': True,
                    'timeout': 60
                }
                if sys.platform == 'win32':
                    kwargs['creationflags'] = 0x08000000  # CREATE_NO_WINDOW
                
                result = subprocess.run(
                    [sys.executable, picker_script, 'file', initial_dir],
                    **kwargs
                )
                
                # 检查stderr中的错误
                if result.stderr:
                    return {'success': False, 'message': '错误: ' + result.stderr.strip()}
                
                output = result.stdout.strip()
                if result.returncode == 0 and output:
                    return {'success': True, 'path': output}
                elif result.returncode == 0:
                    return {'success': False, 'message': '未选择文件'}
                else:
                    return {'success': False, 'message': '选择失败 (code {})'.format(result.returncode)}
        except subprocess.TimeoutExpired:
            return {'success': False, 'message': '选择超时'}
        except Exception as e:
            import traceback
            return {'success': False, 'message': str(e) + '\n' + traceback.format_exc()}
    
    def browse_dir_dialog(self):
        """打开目录选择对话框"""
        try:
            if sys.platform == 'darwin':
                # macOS: 使用osascript
                script = '''
                tell application "System Events"
                    activate
                    set theFolder to choose folder with prompt "选择工作目录"
                    return POSIX path of theFolder
                end tell
                '''
                result = subprocess.run(['osascript', '-e', script],
                                        capture_output=True, text=True, timeout=60)
                if result.returncode == 0 and result.stdout.strip():
                    return {'success': True, 'path': result.stdout.strip().rstrip('/')}
                return {'success': False, 'message': '未选择目录'}
            else:
                # Windows/Linux: 使用独立进程运行tkinter
                script_dir = os.path.dirname(os.path.abspath(__file__))
                picker_script = os.path.join(script_dir, 'file_picker.py')
                
                # Windows上隐藏控制台窗口
                kwargs = {
                    'capture_output': True,
                    'text': True,
                    'timeout': 60
                }
                if sys.platform == 'win32':
                    kwargs['creationflags'] = 0x08000000  # CREATE_NO_WINDOW
                
                result = subprocess.run(
                    [sys.executable, picker_script, 'dir'],
                    **kwargs
                )
                
                # 检查stderr中的错误
                if result.stderr:
                    return {'success': False, 'message': '错误: ' + result.stderr.strip()}
                
                output = result.stdout.strip()
                if result.returncode == 0 and output:
                    return {'success': True, 'path': output}
                elif result.returncode == 0:
                    return {'success': False, 'message': '未选择目录'}
                else:
                    return {'success': False, 'message': '选择失败 (code {})'.format(result.returncode)}
        except subprocess.TimeoutExpired:
            return {'success': False, 'message': '选择超时'}
        except Exception as e:
            import traceback
            return {'success': False, 'message': str(e) + '\n' + traceback.format_exc()}


def main():
    print("=" * 50)
    print("Excel比对工具 - Web界面")
    print("=" * 50)
    print()
    
    if not OPENPYXL_OK:
        print("[警告] 缺少openpyxl库，请运行: pip install openpyxl")
        print()
    
    url = "http://localhost:{}".format(PORT)
    print("启动服务器: {}".format(url))
    print("按 Ctrl+C 停止服务器")
    print()
    
    # 自动打开浏览器
    threading.Timer(1, lambda: webbrowser.open(url)).start()
    
    # 启动服务器
    server = HTTPServer(('localhost', PORT), RequestHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务器已停止")
        server.shutdown()


if __name__ == '__main__':
    main()

