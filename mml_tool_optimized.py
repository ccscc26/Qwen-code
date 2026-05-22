# -*- coding: utf-8 -*-
"""
MML 参数管理工具 - 优化版
支持 Win10 环境，修复乱码和闪退问题
"""

import re
import os
import sys
import json
import time
import logging
import threading
from typing import Dict, List, Set, Optional, Tuple
from collections import defaultdict
from pathlib import Path

# ==========================================
# 1. 环境预检与依赖管理
# ==========================================
def check_dependencies():
    """检查并报告缺失的依赖"""
    missing = []
    try:
        import pandas as pd
    except ImportError:
        missing.append('pandas')
    try:
        import numpy as np
    except ImportError:
        missing.append('numpy')

    if missing:
        print(f"[ERROR] 缺少核心模块：{', '.join(missing)}")
        print(f"请执行：pip install {' '.join(missing)} openpyxl xlsxwriter")
        return False
    return True

if not check_dependencies():
    sys.exit(1)

import pandas as pd
import numpy as np

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    GUI_AVAILABLE = True
except Exception as e:
    print(f"[WARNING] GUI 不可用：{e}")
    GUI_AVAILABLE = False

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

# 配置日志 - 移除 encoding 参数以避免兼容性问题
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# ==========================================
# 2. 常量定义
# ==========================================
class Constants:
    SUPPORTED_ENCODINGS = ['utf-8', 'gbk', 'gb18030', 'utf-16']
    CORE_KEY_COLUMNS = ['网元', 'gNodeB 标识', 'eNodeB 标识', 'NR DU 小区标识', 'NR 小区标识', '本地小区标识']
    TARGET_COLUMNS = [
        'NR DU 小区名称', '小区名称', '网元', 'gNodeB 标识', 'eNodeB 标识',
        'NR DU 小区标识', 'NR 小区标识', '本地小区标识',
        '物理小区标识', '双工模式', '频带', '下行频点', '下行带宽'
    ]
    EXCLUDE_TREE_KEYS = ['网元', 'NR 小区标识', 'NR DU 小区标识', '本地小区标识']
    FILE_TYPES = [("Text Files", "*.txt"), ("All Files", "*.*")]
    EXCEL_INVALID_CHARS = ['\\', '/', '?', '*', ':', '[', ']']
    EXPORT_FORMATS = ["XLSX (Excel)", "CSV"]
    CONFIG_VERSION = "5.0"

    # 大文件优化参数
    MEDIUM_FILE_THRESHOLD = 100 * 1024 * 1024
    STREAM_CHUNK_SIZE = 20 * 1024 * 1024
    BATCH_PROCESS_SIZE = 3000
    MAX_MEMORY_PERCENT = 80.0

REGEX_PATTERNS = {
    'ne_pattern': re.compile(r'\+\+\+\s+(\S+) '),
    'cmd_pattern': re.compile(r'(LST|DSP|MOD|ADD|RMV)\s+(\w+):'),
    'block_split': re.compile(r'(?=\+\+\+\s+\S+)'),
    'header_split': re.compile(r'\s{2,}'),
    'result_count': re.compile(r'\(结果个数|RETCODE'),
}

# ==========================================
# 3. 编码检测与文件读取 (安全增强版)
# ==========================================
def _is_safe_path(filepath: str, allowed_base: Optional[str] = None) -> bool:
    """检查文件路径是否安全，防止目录遍历攻击"""
    try:
        resolved = Path(filepath).resolve()
        
        if '..' in filepath:
            logger.warning(f"检测到目录遍历尝试：{filepath}")
            return False
        
        if allowed_base:
            base_path = Path(allowed_base).resolve()
            if not str(resolved).startswith(str(base_path)):
                logger.warning(f"文件路径超出允许范围：{filepath}")
                return False
        
        if not resolved.is_file():
            logger.warning(f"文件不存在或不是普通文件：{filepath}")
            return False
            
        return True
    except Exception as e:
        logger.error(f"路径验证失败：{e}")
        return False

def _detect_encoding(filepath: str) -> str:
    """检测文件编码，优先检测中文编码"""
    if not _is_safe_path(filepath):
        logger.error(f"不安全的文件路径被拒绝：{filepath}")
        return 'utf-8'
    
    try:
        with open(filepath, 'rb') as f:
            raw = f.read(8192)
        
        # 优先检测含中文的编码
        for enc in ['gb18030', 'gbk', 'utf-8', 'utf-16']:
            try:
                decoded = raw.decode(enc)
                # 如果包含中文字符，认为找到正确编码
                if any('\u4e00' <= c <= '\u9fff' for c in decoded):
                    logger.info(f"检测到编码 (含中文): {enc} for {filepath}")
                    return enc
            except UnicodeDecodeError:
                continue
        
        # 如果没有中文，返回第一个能解码的
        for enc in ['utf-8', 'gb18030', 'gbk', 'utf-16']:
            try:
                raw.decode(enc)
                logger.info(f"检测到编码：{enc} for {filepath}")
                return enc
            except UnicodeDecodeError:
                continue
                
    except FileNotFoundError:
        logger.error(f"文件不存在：{filepath}")
    except PermissionError:
        logger.error(f"无权限读取文件：{filepath}")
    except Exception as e:
        logger.error(f"编码检测异常：{e}")
    
    return 'utf-8'

def read_file_auto_encoding(filename: str) -> str:
    """读取文件并自动检测编码"""
    if not _is_safe_path(filename):
        logger.error(f"不安全的文件路径被拒绝：{filename}")
        return ""
    
    enc = _detect_encoding(filename)
    try:
        with open(filename, 'r', encoding=enc, errors='replace') as f:
            content = f.read()
        logger.info(f"成功读取文件：{filename}, 编码：{enc}")
        return content
    except FileNotFoundError:
        logger.error(f"文件不存在：{filename}")
    except PermissionError:
        logger.error(f"无权限读取文件：{filename}")
    except Exception as e:
        logger.error(f"文件读取异常：{e}")
    return ""

# ==========================================
# 4. MML 解析引擎
# ==========================================
def _parse_blocks_batch(blocks: List[str], raw_data_dict: Dict):
    """解析一批 MML blocks，追加到 raw_data_dict"""
    for block in blocks:
        if '+++' not in block:
            continue

        ne_match = REGEX_PATTERNS['ne_pattern'].search(block)
        ne_name = ne_match.group(1) if ne_match else "Unknown"

        cmd_match = REGEX_PATTERNS['cmd_pattern'].search(block)
        if not cmd_match:
            continue
        cmd_type = f"{cmd_match.group(1)} {cmd_match.group(2)}"

        lines = block.split('\n')
        sep_idx = -1
        for i, line in enumerate(lines):
            if line.strip().startswith('---') and len(line.strip()) >= 3 and 'END' not in line:
                sep_idx = i
                break

        if sep_idx == -1:
            continue

        is_kv = False
        for i in range(sep_idx + 1, min(sep_idx + 5, len(lines))):
            if '=' in lines[i] and not REGEX_PATTERNS['result_count'].search(lines[i]):
                is_kv = True
                break

        data_rows = []
        if is_kv:
            row = {'网元': ne_name}
            for i in range(sep_idx + 1, len(lines)):
                line = lines[i].strip()
                if not line or line.startswith('---'):
                    continue
                if line.startswith('('):
                    break
                if '=' in line:
                    k, v = line.split('=', 1)
                    row[k.strip()] = v.strip()
            if len(row) > 1:
                data_rows.append(row)
        else:
            header_lines = []
            i = sep_idx + 1
            while i < len(lines) and lines[i].strip() != '':
                header_lines.append(lines[i].strip())
                i += 1

            header_str = "     ".join(header_lines)
            headers = REGEX_PATTERNS['header_split'].split(header_str)
            headers = [h.strip() for h in headers if h.strip()]

            if not headers:
                continue

            i += 1
            data_lines = []
            while i < len(lines):
                line_strip = lines[i].strip()
                if line_strip.startswith('(') or line_strip.startswith('---') or \
                   line_strip.startswith('共有') or line_strip.startswith('仍有'):
                    break
                if line_strip:
                    data_lines.append(line_strip)
                i += 1

            num_headers = len(headers)
            if num_headers > 0 and data_lines:
                data_str = "     ".join(data_lines)
                values = REGEX_PATTERNS['header_split'].split(data_str.strip())
                values = [v.strip() for v in values if v.strip()]

                for j in range(0, len(values), num_headers):
                    chunk = values[j:j+num_headers]
                    if len(chunk) > 0:
                        row = {'网元': ne_name}
                        for k, h in enumerate(headers):
                            row[h] = chunk[k] if k < len(chunk) else ''
                        data_rows.append(row)

        if data_rows:
            raw_data_dict.setdefault(cmd_type, []).extend(data_rows)

def parse_mml_file_streaming(filename: str) -> Dict[str, pd.DataFrame]:
    """流式分块解析大文件"""
    enc = _detect_encoding(filename)

    raw_data_dict = defaultdict(list)
    buffer = ""
    block_batch = []

    with open(filename, 'r', encoding=enc, errors='replace') as f:
        while True:
            chunk = f.read(Constants.STREAM_CHUNK_SIZE)
            if not chunk:
                if buffer.strip():
                    last_blocks = REGEX_PATTERNS['block_split'].split(buffer)
                    block_batch.extend([b for b in last_blocks if b.strip()])
                break

            buffer += chunk

            last_marker = buffer.rfind('\n+++ ')
            if last_marker == -1 and buffer.startswith('+++ '):
                last_marker = 0

            if last_marker > 0:
                complete_part = buffer[:last_marker]
                buffer = buffer[last_marker:]

                blocks = REGEX_PATTERNS['block_split'].split(complete_part)
                block_batch.extend([b for b in blocks if b.strip()])

            if len(block_batch) >= Constants.BATCH_PROCESS_SIZE:
                _parse_blocks_batch(block_batch, raw_data_dict)
                block_batch = []

    if block_batch:
        _parse_blocks_batch(block_batch, raw_data_dict)

    table_dict = {}
    for cmd_type, rows in raw_data_dict.items():
        if rows:
            table_dict[cmd_type] = pd.DataFrame(rows)

    return table_dict

def parse_mml_file(filename: str) -> Dict[str, pd.DataFrame]:
    """智能解析入口"""
    file_size = os.path.getsize(filename)

    if file_size >= Constants.MEDIUM_FILE_THRESHOLD:
        logger.info(f"启用流式解析：{os.path.basename(filename)} ({file_size/1024/1024:.1f}MB)")
        return parse_mml_file_streaming(filename)
    else:
        content = read_file_auto_encoding(filename)
        if not content:
            return {}

        raw_data_dict = {}
        blocks = REGEX_PATTERNS['block_split'].split(content)
        _parse_blocks_batch(blocks, raw_data_dict)

        table_dict = {}
        for cmd_type, rows in raw_data_dict.items():
            if rows:
                table_dict[cmd_type] = pd.DataFrame(rows)
        return table_dict

# ==========================================
# 5. 数据过滤与融合引擎
# ==========================================
def apply_row_filters(table_dict, group_filters, fid):
    if fid not in group_filters:
        return table_dict

    filtered_dict = {}
    for cmd, df in table_dict.items():
        if cmd in group_filters.get(fid, {}):
            new_df = df.copy()
            for field, values in group_filters[fid][cmd].items():
                if values and field in new_df.columns:
                    mask = new_df[field].astype(str).isin([str(v) for v in values])
                    new_df = new_df[mask]
            filtered_dict[cmd] = new_df
        else:
            filtered_dict[cmd] = df
    return filtered_dict

def align_tables_by_cell(table_dict, selected_fields=None):
    """核心融合引擎"""
    if not table_dict:
        return pd.DataFrame()

    main_cmd = None
    for candidate in ['LST NRDUCELL', 'LST NRCELL', 'LST CELL']:
        if candidate in table_dict:
            main_cmd = candidate
            break

    if not main_cmd:
        for cmd in table_dict:
            cols = table_dict[cmd].columns
            if any(k in cols for k in ['NR DU 小区标识', 'NR 小区标识', '本地小区标识']):
                main_cmd = cmd
                break
        else:
            main_cmd = list(table_dict.keys())[0]

    main_df = table_dict[main_cmd].copy()

    main_cell_key = None
    for k in ['NR DU 小区标识', 'NR 小区标识', '本地小区标识']:
        if k in main_df.columns:
            main_cell_key = k
            break

    core_keys = Constants.CORE_KEY_COLUMNS

    if selected_fields and main_cmd in selected_fields and selected_fields[main_cmd]:
        keep_cols = [c for c in core_keys if c in main_df.columns]
        extra_cols = [c for c in selected_fields[main_cmd] if c in main_df.columns]
        keep_cols += extra_cols
        keep_cols = list(dict.fromkeys(keep_cols))
        main_df = main_df[keep_cols]

    if '网元' not in main_df.columns:
        return pd.DataFrame()

    main_df = main_df.drop_duplicates()

    try:
        if not main_cell_key:
            for cmd, df in table_dict.items():
                if cmd == main_cmd:
                    continue
                sub_df = df.copy()
                if selected_fields and cmd in selected_fields and selected_fields[cmd]:
                    sub_cols = [c for c in core_keys if c in sub_df.columns]
                    sub_cols += [c for c in selected_fields[cmd] if c in sub_df.columns]
                    sub_df = sub_df[list(dict.fromkeys(sub_cols))]

                sub_df = sub_df.drop_duplicates(subset=['网元'])
                overlap = set(main_df.columns) & set(sub_df.columns) - set(core_keys)
                if overlap:
                    rename_dict = {}
                    for c in overlap:
                        if c in sub_df.columns:
                            rename_dict[c] = f"{c}_{cmd.split()[-1]}"
                    if rename_dict:
                        sub_df = sub_df.rename(columns=rename_dict)
                main_df = main_df.merge(sub_df, on=['网元'], how='left')
            return main_df

        cell_max_rows = main_df.groupby(['网元', main_cell_key]).size().to_dict()
        processed = {}

        for cmd, df in table_dict.items():
            if cmd == main_cmd:
                continue

            sub_df = df.copy()

            if selected_fields:
                if cmd not in selected_fields or not selected_fields[cmd]:
                    continue
                sub_cols = [c for c in core_keys if c in sub_df.columns]
                sub_cols += [c for c in selected_fields[cmd] if c in sub_df.columns]
                sub_df = sub_df[list(dict.fromkeys(sub_cols))]

            sub_df = sub_df.drop_duplicates()

            sub_key = None
            for k in ['NR 小区标识', 'NR DU 小区标识', '本地小区标识']:
                if k in sub_df.columns:
                    sub_key = k
                    break

            if not sub_key:
                processed[cmd] = (sub_df, False)
                continue

            if sub_key != main_cell_key:
                sub_df = sub_df.rename(columns={sub_key: main_cell_key})

            sub_df['_seq_idx'] = sub_df.groupby(['网元', main_cell_key]).cumcount()

            for (ne, cid), count in sub_df.groupby(['网元', main_cell_key]).size().items():
                cell_max_rows[(ne, cid)] = max(cell_max_rows.get((ne, cid), 1), count)

            overlap = set(main_df.columns) & set(sub_df.columns) - set(core_keys) - {'_seq_idx'}
            if overlap:
                rename_dict = {}
                for c in overlap:
                    if c in sub_df.columns:
                        rename_dict[c] = f"{c}_{cmd.split()[-1]}"
                if rename_dict:
                    sub_df = sub_df.rename(columns=rename_dict)

            processed[cmd] = (sub_df, True)

        exp_records = [
            {'网元': ne, main_cell_key: cid, '_seq_idx': r}
            for (ne, cid), max_r in cell_max_rows.items()
            for r in range(max_r)
        ]

        if not exp_records:
            return main_df

        exp_df = pd.DataFrame(exp_records)
        ext_main = exp_df.merge(main_df, on=['网元', main_cell_key], how='left')

        for cmd, (sub_df, is_cell) in processed.items():
            if is_cell:
                ext_main = ext_main.merge(
                    sub_df,
                    on=['网元', main_cell_key, '_seq_idx'],
                    how='left'
                )
            else:
                ext_main = ext_main.merge(
                    sub_df.drop_duplicates(subset=['网元']),
                    on=['网元'],
                    how='left'
                )

        if '_seq_idx' in ext_main.columns:
            ext_main.drop('_seq_idx', axis=1, inplace=True)

        return ext_main

    except Exception as e:
        logger.error(f"表融合失败：{e}")
        import traceback
        traceback.print_exc()
        return main_df

# ==========================================
# 6. 导出功能
# ==========================================
def export_to_multi_sheets(data_dict: Dict[str, pd.DataFrame], output_path: str) -> Tuple[bool, str]:
    """导出多 Sheet Excel"""
    try:
        if output_path.endswith('.xlsx'):
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in data_dict.items():
                    safe_name = sheet_name.replace('/', '_').replace('\\', '_')[:31]
                    df.to_excel(writer, sheet_name=safe_name, index=False)
            return True, f"成功导出到：{output_path}"
        elif output_path.endswith('.csv'):
            for name, df in data_dict.items():
                csv_path = output_path.replace('.csv', f'_{name}.csv')
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            return True, f"成功导出 CSV 文件到：{output_path}"
        else:
            return False, "不支持的文件格式"
    except Exception as e:
        logger.error(f"导出失败：{e}")
        return False, f"导出失败：{str(e)}"

# ==========================================
# 7. GUI 主界面 (Win10 优化版)
# ==========================================
class MMLToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MML 参数管理工具 v5.0 - Win10 优化版")
        self.root.geometry("1000x700")
        
        # 设置 Windows 风格字体 (解决中文乱码)
        self.setup_fonts()
        
        self.files_data = []
        self.group_filters = {}
        self.selected_fields = {}
        self.config_file = "mml_tool_config.json"
        
        self.create_widgets()
        self.load_config()

    def setup_fonts(self):
        """设置适合 Windows 的中文字体"""
        try:
            # Windows 常用中文字体
            if sys.platform == 'win32':
                default_font = ('Microsoft YaHei', 10)  # 微软雅黑
                tree_font = ('Microsoft YaHei', 9)
                button_font = ('Microsoft YaHei', 9, 'bold')
            else:
                default_font = ('Arial Unicode MS', 10)
                tree_font = ('Arial Unicode MS', 9)
                button_font = ('Arial Unicode MS', 9, 'bold')
            
            self.root.option_add('*Font', default_font[0])
            self.default_font = default_font
            self.tree_font = tree_font
            self.button_font = button_font
        except Exception as e:
            logger.warning(f"字体设置失败：{e}, 使用默认字体")
            self.default_font = ('TkDefaultFont', 10)
            self.tree_font = ('TkTextFont', 9)
            self.button_font = ('TkDefaultFont', 9, 'bold')

    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # 顶部按钮区
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(btn_frame, text="添加 TXT 文件", command=self.add_txt_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="刷新显示", command=self.refresh_tree).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="融合数据", command=self.merge_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="导出 Excel", command=self.export_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="自校验", command=self.self_check).pack(side=tk.LEFT, padx=5)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        # 树形视图框架
        tree_frame = ttk.Frame(main_frame)
        tree_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        # 树形控件
        self.tree = ttk.Treeview(
            tree_frame, 
            columns=('文件', '键', '值'),
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        self.tree.heading('#0', text='层级结构')
        self.tree.heading('文件', text='文件')
        self.tree.heading('键', text='键')
        self.tree.heading('值', text='值')
        
        self.tree.column('#0', width=300, minwidth=200)
        self.tree.column('文件', width=150, minwidth=100)
        self.tree.column('键', width=150, minwidth=100)
        self.tree.column('值', width=200, minwidth=150)
        
        # 布局
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

    def add_txt_file(self):
        """添加 TXT 文件"""
        file_path = filedialog.askopenfilename(
            title="选择 TXT 参数文件",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            if not _is_safe_path(file_path):
                messagebox.showerror("错误", "不安全的文件路径!")
                return
            
            # 检测编码并读取
            encoding = _detect_encoding(file_path)
            logger.info(f"添加文件：{file_path}, 编码：{encoding}")
            
            content = read_file_auto_encoding(file_path)
            if not content:
                messagebox.showwarning("警告", "无法读取文件内容")
                return
            
            # 解析 MML
            table_dict = parse_mml_file(file_path)
            
            if not table_dict:
                messagebox.showwarning("警告", "未解析到有效数据")
                return
            
            # 存储文件数据
            file_id = len(self.files_data)
            self.files_data.append({
                'id': file_id,
                'path': file_path,
                'name': os.path.basename(file_path),
                'encoding': encoding,
                'table_dict': table_dict
            })
            
            # 添加到树形视图
            self.add_file_to_tree(file_id)
            
            self.status_var.set(f"已添加：{os.path.basename(file_path)}")
            
        except Exception as e:
            logger.error(f"添加文件失败：{e}")
            messagebox.showerror("错误", f"添加文件失败:\n{str(e)}")

    def add_file_to_tree(self, file_id: int):
        """将文件数据添加到树形视图"""
        file_info = self.files_data[file_id]
        file_name = file_info['name']
        table_dict = file_info['table_dict']
        encoding = file_info['encoding']
        
        # 添加文件节点 (显示编码信息)
        file_node = self.tree.insert(
            "", "end",
            text=f"📄 {file_name} [{encoding}]",
            values=(file_name, "", "")
        )
        
        # 添加命令表节点
        for cmd_type, df in table_dict.items():
            cmd_node = self.tree.insert(
                file_node, "end",
                text=f"📋 {cmd_type} ({len(df)}行)",
                values=(file_name, cmd_type, f"{len(df)}行")
            )
            
            # 添加列头信息
            for col in df.columns:
                self.tree.insert(
                    cmd_node, "end",
                    text=f"  ├─ {col}",
                    values=(file_name, col, "列")
                )
                
                # 添加前 5 行示例数据
                for idx in range(min(5, len(df))):
                    val = str(df.iloc[idx][col])
                    if len(val) > 50:
                        val = val[:47] + "..."
                    self.tree.insert(
                        cmd_node, "end",
                        text=f"  │  [{idx}] {val}",
                        values=(file_name, f"{col}[{idx}]", val)
                    )
        
        # 自动展开
        self.tree.item(file_node, open=True)

    def refresh_tree(self):
        """刷新树形视图"""
        # 清空
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 重新添加
        for file_info in self.files_data:
            self.add_file_to_tree(file_info['id'])
        
        self.status_var.set("显示已刷新")

    def merge_data(self):
        """融合数据"""
        if not self.files_data:
            messagebox.showwarning("警告", "请先添加文件")
            return
        
        try:
            self.status_var.set("正在融合数据...")
            self.root.update_idletasks()
            
            all_table_dicts = []
            for file_info in self.files_data:
                filtered = apply_row_filters(
                    file_info['table_dict'],
                    self.group_filters,
                    file_info['id']
                )
                all_table_dicts.append(filtered)
            
            # 合并所有表
            merged_dict = defaultdict(list)
            for table_dict in all_table_dicts:
                for cmd, df in table_dict.items():
                    if not df.empty:
                        merged_dict[cmd].append(df)
            
            final_dict = {}
            for cmd, dfs in merged_dict.items():
                if dfs:
                    final_dict[cmd] = pd.concat(dfs, ignore_index=True).drop_duplicates()
            
            if not final_dict:
                messagebox.showwarning("警告", "未融合到任何数据")
                self.status_var.set("就绪")
                return
            
            # 显示融合结果
            result_window = tk.Toplevel(self.root)
            result_window.title("融合结果")
            result_window.geometry("800x600")
            
            text = tk.Text(result_window, wrap=tk.WORD)
            text.pack(fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(text, command=text.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            text.config(yscrollcommand=scrollbar.set)
            
            for cmd, df in final_dict.items():
                text.insert(tk.END, f"=== {cmd} ({len(df)}行) ===\n")
                text.insert(tk.END, df.to_string() + "\n\n")
            
            text.config(state=tk.DISABLED)
            
            self.status_var.set(f"融合完成：{len(final_dict)}个表")
            
        except Exception as e:
            logger.error(f"融合失败：{e}")
            messagebox.showerror("错误", f"融合失败:\n{str(e)}")
            self.status_var.set("就绪")

    def export_excel(self):
        """导出 Excel"""
        if not self.files_data:
            messagebox.showwarning("警告", "请先添加文件")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
        )
        
        if not file_path:
            return
        
        try:
            self.status_var.set("正在导出...")
            self.root.update_idletasks()
            
            merged_dict = defaultdict(list)
            for file_info in self.files_data:
                filtered = apply_row_filters(
                    file_info['table_dict'],
                    self.group_filters,
                    file_info['id']
                )
                for cmd, df in filtered.items():
                    if not df.empty:
                        merged_dict[cmd].append(df)
            
            final_dict = {}
            for cmd, dfs in merged_dict.items():
                if dfs:
                    final_dict[cmd] = pd.concat(dfs, ignore_index=True).drop_duplicates()
            
            success, msg = export_to_multi_sheets(final_dict, file_path)
            
            if success:
                messagebox.showinfo("成功", msg)
            else:
                messagebox.showerror("错误", msg)
            
            self.status_var.set("就绪")
            
        except Exception as e:
            logger.error(f"导出失败：{e}")
            messagebox.showerror("错误", f"导出失败:\n{str(e)}")
            self.status_var.set("就绪")

    def self_check(self):
        """自校验功能"""
        try:
            self.status_var.set("正在自校验...")
            self.root.update_idletasks()
            
            results = []
            
            # 检查依赖
            results.append("✓ 依赖检查：pandas, numpy 已加载")
            
            # 检查文件
            if not self.files_data:
                results.append("⚠ 未添加任何文件")
            else:
                for f in self.files_data:
                    results.append(f"✓ 文件：{f['name']} (编码:{f['encoding']}, {len(f['table_dict'])}个表)")
            
            # 检查 GUI
            results.append(f"✓ GUI: tkinter 正常 (平台：{sys.platform})")
            
            # 显示结果
            result_window = tk.Toplevel(self.root)
            result_window.title("自校验结果")
            result_window.geometry("500x400")
            
            text = tk.Text(result_window, wrap=tk.WORD)
            text.pack(fill=tk.BOTH, expand=True)
            
            for line in results:
                text.insert(tk.END, line + "\n")
            
            text.config(state=tk.DISABLED)
            
            self.status_var.set("自校验完成")
            
        except Exception as e:
            logger.error(f"自校验失败：{e}")
            messagebox.showerror("错误", f"自校验失败:\n{str(e)}")
            self.status_var.set("就绪")

    def save_config(self):
        """保存配置"""
        try:
            config = {
                'version': Constants.CONFIG_VERSION,
                'files': [
                    {
                        'path': f['path'],
                        'encoding': f['encoding']
                    }
                    for f in self.files_data
                ]
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("成功", f"配置已保存到：{self.config_file}")
            self.status_var.set("配置已保存")
            
        except Exception as e:
            logger.error(f"保存配置失败：{e}")
            messagebox.showerror("错误", f"保存配置失败:\n{str(e)}")

    def load_config(self):
        """加载配置"""
        if not os.path.exists(self.config_file):
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            for file_info in config.get('files', []):
                path = file_info.get('path')
                if path and os.path.exists(path):
                    encoding = file_info.get('encoding', 'utf-8')
                    table_dict = parse_mml_file(path)
                    
                    file_id = len(self.files_data)
                    self.files_data.append({
                        'id': file_id,
                        'path': path,
                        'name': os.path.basename(path),
                        'encoding': encoding,
                        'table_dict': table_dict
                    })
                    self.add_file_to_tree(file_id)
            
            self.status_var.set(f"已加载配置：{len(self.files_data)}个文件")
            
        except Exception as e:
            logger.error(f"加载配置失败：{e}")

# ==========================================
# 8. 主入口
# ==========================================
def main():
    """主函数"""
    if not GUI_AVAILABLE:
        print("[ERROR] GUI 不可用，请确保在图形界面环境下运行")
        print("Windows: 确保安装了 Python tkinter")
        print("Linux: sudo apt-get install python3-tk")
        sys.exit(1)
    
    try:
        root = tk.Tk()
        
        # 设置 DPI 感知 (Windows)
        if sys.platform == 'win32':
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass
        
        app = MMLToolGUI(root)
        root.mainloop()
        
    except tk.TclError as e:
        if "no display" in str(e).lower() or "display" in str(e).lower():
            print("[ERROR] 无法连接到显示服务器")
            print("Windows: 请确保在桌面环境下运行")
            print("Linux: 请设置 DISPLAY 环境变量或使用 Xvfb")
        else:
            print(f"[ERROR] Tkinter 错误：{e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"程序启动失败：{e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
