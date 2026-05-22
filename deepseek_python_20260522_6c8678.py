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
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("环境错误", f"缺少核心模块：{', '.join(missing)}\n"
                                f"请执行：pip install {' '.join(missing)} openpyxl xlsxwriter")
            root.destroy()
        except tk.TclError:
            # 无 GUI 环境，使用命令行提示
            print(f"[ERROR] 缺少核心模块：{', '.join(missing)}")
            print(f"请执行：pip install {' '.join(missing)} openpyxl xlsxwriter")
        except Exception as e:
            print(f"[ERROR] 缺少核心模块：{', '.join(missing)} (初始化错误：{e})")
        sys.exit(1)

check_dependencies()

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

logging.basicConfig(level=logging.INFO, format='%(asctime)s - [%(levelname)s] - %(message)s', encoding='utf-8')
logger = logging.getLogger(__name__)

# ==========================================
# 2. 常量定义
# ==========================================
class Constants:
    SUPPORTED_ENCODINGS = ['utf-8', 'gbk', 'gb18030', 'utf-16']
    CORE_KEY_COLUMNS = ['网元', 'gNodeB标识', 'eNodeB标识', 'NR DU小区标识', 'NR小区标识', '本地小区标识']
    TARGET_COLUMNS = [
        'NR DU小区名称', '小区名称', '网元', 'gNodeB标识', 'eNodeB标识',
        'NR DU小区标识', 'NR小区标识', '本地小区标识',
        '物理小区标识', '双工模式', '频带', '下行频点', '下行带宽'
    ]
    EXCLUDE_TREE_KEYS = ['网元', 'NR小区标识', 'NR DU小区标识', '本地小区标识']
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
        # 解析绝对路径
        resolved = Path(filepath).resolve()
        
        # 检查是否包含目录遍历模式
        if '..' in filepath:
            logger.warning(f"检测到目录遍历尝试：{filepath}")
            return False
        
        # 如果指定了允许的基础目录，检查是否在该目录下
        if allowed_base:
            base_path = Path(allowed_base).resolve()
            if not str(resolved).startswith(str(base_path)):
                logger.warning(f"文件路径超出允许范围：{filepath}")
                return False
        
        # 检查是否为有效文件
        if not resolved.is_file():
            logger.warning(f"文件不存在或不是普通文件：{filepath}")
            return False
            
        return True
    except Exception as e:
        logger.error(f"路径验证失败：{e}")
        return False

def _detect_encoding(filepath: str) -> str:
    """检测文件编码，增加路径安全检查"""
    # 安全校验
    if not _is_safe_path(filepath):
        logger.error(f"不安全的文件路径被拒绝：{filepath}")
        return 'utf-8'
    
    try:
        with open(filepath, 'rb') as f:
            raw = f.read(8192)
        for enc in Constants.SUPPORTED_ENCODINGS:
            try:
                decoded = raw.decode(enc)
                if any('\u4e00' <= c <= '\u9fff' for c in decoded):
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
    """读取文件并自动检测编码，增加路径安全检查"""
    # 安全校验
    if not _is_safe_path(filename):
        logger.error(f"不安全的文件路径被拒绝：{filename}")
        return ""
    
    enc = _detect_encoding(filename)
    try:
        with open(filename, 'r', encoding=enc, errors='replace') as f:
            return f.read()
    except FileNotFoundError:
        logger.error(f"文件不存在：{filename}")
    except PermissionError:
        logger.error(f"无权限读取文件：{filename}")
    except Exception as e:
        logger.error(f"文件读取异常：{e}")
    return ""

# ==========================================
# 4. MML解析引擎
# ==========================================
def _parse_blocks_batch(blocks: List[str], raw_data_dict: Dict):
    """解析一批MML blocks，追加到raw_data_dict"""
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
        logger.info(f"启用流式解析: {os.path.basename(filename)} ({file_size/1024/1024:.1f}MB)")
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
# 5. 数据过滤与融合引擎（修复版）
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
    """
    核心融合引擎（修复版）
    修复：Columns must be same length as key
    根因：子表筛选/重命名后列数不一致
    """
    if not table_dict:
        return pd.DataFrame()

    # 确定主表
    main_cmd = None
    for candidate in ['LST NRDUCELL', 'LST NRCELL', 'LST CELL']:
        if candidate in table_dict:
            main_cmd = candidate
            break

    if not main_cmd:
        for cmd in table_dict:
            cols = table_dict[cmd].columns
            if any(k in cols for k in ['NR DU小区标识', 'NR小区标识', '本地小区标识']):
                main_cmd = cmd
                break
        else:
            main_cmd = list(table_dict.keys())[0]

    main_df = table_dict[main_cmd].copy()

    # 确定主小区键
    main_cell_key = None
    for k in ['NR DU小区标识', 'NR小区标识', '本地小区标识']:
        if k in main_df.columns:
            main_cell_key = k
            break

    core_keys = Constants.CORE_KEY_COLUMNS

    # 字段筛选（安全版本）
    if selected_fields and main_cmd in selected_fields and selected_fields[main_cmd]:
        keep_cols = [c for c in core_keys if c in main_df.columns]
        extra_cols = [c for c in selected_fields[main_cmd] if c in main_df.columns]
        keep_cols += extra_cols
        keep_cols = list(dict.fromkeys(keep_cols))  # 去重保序
        main_df = main_df[keep_cols]

    if '网元' not in main_df.columns:
        return pd.DataFrame()

    main_df = main_df.drop_duplicates()

    try:
        # 非小区级别合并
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
                # 安全重命名：只重命名存在于sub_df中的列
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

        # 小区级别合并
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
            for k in ['NR小区标识', 'NR DU小区标识', '本地小区标识']:
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

            # 安全重命名
            overlap = set(main_df.columns) & set(sub_df.columns) - set(core_keys) - {'_seq_idx'}
            if overlap:
                rename_dict = {}
                for c in overlap:
                    if c in sub_df.columns:
                        rename_dict[c] = f"{c}_{cmd.split()[-1]}"
                if rename_dict:
                    sub_df = sub_df.rename(columns=rename_dict)

            processed[cmd] = (sub_df, True)

        # 构建扩展主表
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
                ext_main = ext_main.merge(sub_df, on=['网元', main_cell_key, '_seq_idx'], how='left')
            else:
                sub_u = sub_df.drop_duplicates(subset=['网元'])
                overlap = set(ext_main.columns) & set(sub_u.columns) - set(core_keys)
                if overlap:
                    rename_dict = {}
                    for c in overlap:
                        if c in sub_u.columns:
                            rename_dict[c] = f"{c}_{cmd.split()[-1]}"
                    if rename_dict:
                        sub_u = sub_u.rename(columns=rename_dict)
                ext_main = ext_main.merge(sub_u, on=['网元'], how='left')

        ext_main = ext_main.drop(columns=['_seq_idx'], errors='ignore')

        if not ext_main.empty:
            cols = ext_main.columns.tolist()
            tmp = ext_main.replace('', np.nan)
            fill_cols = [c for c in tmp.columns if c not in ['网元', main_cell_key]]
            if fill_cols:
                tmp[fill_cols] = tmp.groupby(['网元', main_cell_key])[fill_cols].ffill()
                tmp[fill_cols] = tmp.groupby(['网元', main_cell_key])[fill_cols].bfill()
            ext_main = tmp.fillna('').reindex(columns=cols).reset_index(drop=True)

        return ext_main

    except Exception as e:
        logger.error(f"融合引擎降级恢复: {e}")
        return main_df

# ==========================================
# 6. Excel导出
# ==========================================
def export_to_xlsx_format(df, output_path):
    if df.empty:
        return False, "没有数据可导出"

    final_cols = [c for c in Constants.TARGET_COLUMNS if c in df.columns]
    for col in df.columns:
        if col not in final_cols:
            final_cols.append(col)

    df = df[final_cols]

    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='参数对齐大表', na_rep='')
            worksheet = writer.sheets['参数对齐大表']
            for idx, col in enumerate(df.columns):
                series_sample = df[col].head(100).astype(str)
                max_len = max(series_sample.map(len).max(), len(str(col))) + 3
                worksheet.set_column(idx, idx, min(max_len, 35))
        return True, f"成功导出 {len(df)} 行数据 (XLSX格式)。"
    except Exception:
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='参数对齐大表', na_rep='')
            return True, f"成功导出 {len(df)} 行数据 (XLSX兼容模式)。"
        except Exception as err:
            return False, str(err)

def export_to_multi_sheets(cmd_dfs, output_path):
    if not cmd_dfs:
        return False, "没有数据可供导出"

    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            for cmd_name, df in cmd_dfs.items():
                df_copy = df.copy()
                front_cols = [
                    'NR DU小区名称', '小区名称', '网元', 'gNodeB标识', 'eNodeB标识',
                    'NR DU小区标识', 'NR小区标识', '本地小区标识'
                ]
                ordered_cols = [c for c in front_cols if c in df_copy.columns] + \
                              [c for c in df_copy.columns if c not in front_cols]
                df_copy = df_copy[ordered_cols]

                clean_name = cmd_name
                for char in Constants.EXCEL_INVALID_CHARS:
                    clean_name = clean_name.replace(char, '_')
                clean_name = clean_name[:31] if clean_name else '参数对齐大表'

                df_copy.to_excel(writer, index=False, sheet_name=clean_name, na_rep='')
                worksheet = writer.sheets[clean_name]

                for idx, col in enumerate(df_copy.columns):
                    series_sample = df_copy[col].head(100).astype(str)
                    max_len = max(series_sample.map(len).max(), len(str(col))) + 3
                    worksheet.set_column(idx, idx, min(max_len, 35))
        return True, f"成功生成 {len(cmd_dfs)} 个工作表"
    except Exception as err:
        return False, str(err)

# ==========================================
# 7. GUI应用层
# ==========================================
class MMLToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("4/5G MML 批量导出汇总工具 v5.0")
        self.root.geometry("1350x850")

        self.style = ttk.Style()
        # 使用跨平台的中文字体配置，修复乱码问题
        if sys.platform == 'win32':
            font_family = 'Microsoft YaHei UI'
        elif sys.platform == 'darwin':
            font_family = 'PingFang SC'
        else:  # Linux
            font_family = 'WenQuanYi Micro Hei'
        
        self.style.configure('.', font=(font_family, 10))
        self.style.configure('Treeview', font=(font_family, 10), rowheight=26)
        self.style.configure('Treeview.Heading', font=(font_family, 10, 'bold'))
        self.style.configure('Action.TButton', font=(font_family, 10, 'bold'))
        self.style.configure('NewAction.TButton', font=(font_family, 10, 'bold'), foreground='blue')
        self.style.configure('MatchAction.TButton', font=(font_family, 10, 'bold'), foreground='purple')

        self.files_data = []
        self.selected_fields = {}
        self.group_filters = {}
        self.file_id_counter = 0

        self._create_widgets()

    def _create_widgets(self):
        """创建GUI界面"""
        top = ttk.Frame(self.root, padding="10")
        top.pack(fill=tk.X)

        ttk.Button(top, text="1. 批量加载MML", command=self.load_files).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="2. 导入配置(.json)", command=self.load_config).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="同步到所有文件", command=self.apply_config_to_all).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="3. 导出配置", command=self.save_config).pack(side=tk.LEFT, padx=4)

        ttk.Button(top, text="5. 命令全量导出", style='NewAction.TButton',
                  command=self.export_all_per_sheet).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="4. 合并导出", style='Action.TButton',
                  command=self.export_all_merged).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="匹配5G信息", style='MatchAction.TButton',
                  command=self.match_5g_and_export).pack(side=tk.RIGHT, padx=4)

        main = ttk.Frame(self.root, padding="10")
        main.pack(fill=tk.BOTH, expand=True)

        left = ttk.LabelFrame(main, text="已加载的 MML 数据源列表 (支持右键删除)", padding="5")
        left.pack(side=tk.LEFT, fill=tk.BOTH, padx=5)

        self.file_lb = tk.Listbox(left, width=38, height=35, selectmode=tk.BROWSE)
        self.file_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="移除此数据源", command=self.delete_selected_file)
        self.file_lb.bind("<Button-3>", self._show_context_menu)

        sb = ttk.Scrollbar(left, orient="vertical", command=self.file_lb.yview)
        sb.pack(side=tk.RIGHT, fill="y")
        self.file_lb.configure(yscrollcommand=sb.set)
        self.file_lb.bind('<<ListboxSelect>>', self.on_file_select)

        right = ttk.LabelFrame(main, text="过滤面板 (双击切换勾选，单击展开/折叠)", padding="5")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        ctrl_bar = ttk.Frame(right)
        ctrl_bar.pack(fill=tk.X, pady=2, padx=2)

        ttk.Button(ctrl_bar, text="全选字段", command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl_bar, text="清空字段", command=self.deselect_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl_bar, text="[-] 收起", width=15, command=self.collapse_all_nodes).pack(side=tk.RIGHT, padx=2)
        ttk.Button(ctrl_bar, text="[+] 展开", width=15, command=self.expand_all_nodes).pack(side=tk.RIGHT, padx=2)

        self.param_tree = ttk.Treeview(right, show="tree headings")
        self.param_tree.heading("#0", text="MML命令 / 参数字段名", anchor="w")
        self.param_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tree_sb = ttk.Scrollbar(right, orient="vertical", command=self.param_tree.yview)
        tree_sb.pack(side=tk.RIGHT, fill="y")
        self.param_tree.configure(yscrollcommand=tree_sb.set)

        self.param_tree.bind('<Double-1>', self._on_tree_double_click)

        self.status = ttk.Label(self.root, text="就绪。请加载MML文本数据...", relief=tk.SUNKEN)
        self.status.pack(fill=tk.X, side=tk.BOTTOM)

    def expand_all_nodes(self):
        for child_id in self.param_tree.get_children():
            self.param_tree.item(child_id, open=True)

    def collapse_all_nodes(self):
        for child_id in self.param_tree.get_children():
            self.param_tree.item(child_id, open=False)

    def load_files(self):
        paths = filedialog.askopenfilenames(filetypes=Constants.FILE_TYPES)
        if not paths:
            return

        def worker():
            success_count = 0
            for path in paths:
                if any(f['path'] == path for f in self.files_data):
                    continue

                file_size = os.path.getsize(path)
                size_str = f"{file_size/1024/1024:.1f}MB" if file_size > 1024*1024 else f"{file_size/1024:.1f}KB"
                self.root.after(0, lambda p=path, s=size_str: self.status.config(
                    text=f"正在解析: {os.path.basename(p)} ({s}) ..."))
                self.root.after(0, self.root.update_idletasks)

                try:
                    start_time = time.time()
                    table_dict = parse_mml_file(path)
                    elapsed = time.time() - start_time

                    if not table_dict:
                        self.root.after(0, lambda p=path:
                                       messagebox.showwarning("解析警告",
                                           f"文件 {os.path.basename(p)} 未解析到有效数据"))
                        continue

                    fid = self.file_id_counter
                    self.file_id_counter += 1

                    self.files_data.append({'id': fid, 'path': path, 'table_dict': table_dict})
                    self.root.after(0, lambda f=fid, p=path, t=table_dict, e=elapsed:
                                   self.file_lb.insert(tk.END,
                                       f"[{len(self.files_data)}] {os.path.basename(p)} "
                                       f"({len(t)}个命令, {e:.1f}s)"))
                    success_count += 1
                except Exception as e:
                    self.root.after(0, lambda p=path, err=str(e):
                                   messagebox.showerror("解析失败",
                                       f"文件 {os.path.basename(p)} 解析异常:\n{err}"))

            if success_count > 0:
                last_idx = len(self.files_data) - 1
                self.root.after(0, lambda: self.file_lb.selection_set(last_idx))
                self.root.after(0, lambda: self.show_fields_tree(last_idx))
                self.root.after(0, lambda: self.status.config(
                    text=f"成功批量导入 {success_count} 个MML报文文件。"))
            else:
                self.root.after(0, lambda: self.status.config(text="就绪。"))

        threading.Thread(target=worker, daemon=True).start()

    def on_file_select(self, event):
        sel = self.file_lb.curselection()
        if sel:
            self.show_fields_tree(sel[0])

    def _show_context_menu(self, event):
        idx = self.file_lb.nearest(event.y)
        if idx >= 0:
            bbox = self.file_lb.bbox(idx)
            if bbox and event.y <= bbox[1] + bbox[3]:
                self.file_lb.selection_clear(0, tk.END)
                self.file_lb.selection_set(idx)
                self.file_lb.activate(idx)
                self.show_fields_tree(idx)
                self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_file(self):
        sel = self.file_lb.curselection()
        if not sel:
            return
        idx = sel[0]

        file_info = self.files_data[idx]
        fid = file_info['id']
        filename = os.path.basename(file_info['path'])

        if messagebox.askyesno("确认移除",
                               f"确定从工具中移除该 MML 数据源吗？\n\n"
                               f"文件: {filename}\n"
                               f"提醒: 该文件的字段勾选及过滤条件将会被同步销毁。"):
            if fid in self.selected_fields:
                del self.selected_fields[fid]
            if fid in self.group_filters:
                del self.group_filters[fid]

            self.files_data.pop(idx)
            self.file_lb.delete(idx)

            if self.files_data:
                new_idx = min(idx, len(self.files_data) - 1)
                self.file_lb.selection_set(new_idx)
                self.show_fields_tree(new_idx)
                self.status.config(text=f"已成功移除数据源: {filename}")
            else:
                self.param_tree.delete(*self.param_tree.get_children())
                self.status.config(text="数据源已全部清空。")

    def show_fields_tree(self, idx):
        """渲染属性过滤树（修复中文乱码）"""
        old_states = {}
        for child_id in self.param_tree.get_children():
            old_states[child_id] = self.param_tree.item(child_id, "open")
            for sub_child_id in self.param_tree.get_children(child_id):
                old_states[sub_child_id] = self.param_tree.item(sub_child_id, "open")

        self.param_tree.delete(*self.param_tree.get_children())

        if idx >= len(self.files_data):
            return

        file_info = self.files_data[idx]
        fid = file_info['id']
        table_dict = file_info['table_dict']

        self.selected_fields.setdefault(fid, {})
        self.group_filters.setdefault(fid, {})

        for cmd in sorted(table_dict.keys()):
            df = table_dict[cmd]
            if df.empty:
                continue

            fields = [c for c in df.columns if c not in Constants.EXCLUDE_TREE_KEYS]
            self.selected_fields[fid].setdefault(cmd, set())
            current_selected = self.selected_fields[fid][cmd]

            if not fields:
                parent_text = f"[+] {cmd} (no fields)"
            elif len(current_selected) == len(fields):
                parent_text = f"[v] {cmd} ({len(fields)} fields)"
            elif len(current_selected) > 0:
                parent_text = f"[*] {cmd} ({len(current_selected)}/{len(fields)})"
            else:
                parent_text = f"[ ] {cmd} ({len(fields)} fields)"

            parent_node = self.param_tree.insert(
                "", "end", iid=cmd, text=parent_text,
                open=old_states.get(cmd, True)
            )

            for f in fields:
                is_checked = f in current_selected
                prefix = "[v] " if is_checked else "[ ] "
                child_iid = f"{cmd}|||{f}"

                self.param_tree.insert(parent_node, "end", iid=child_iid,
                                      text=f"{prefix}{str(f)}")

                if any(kw in f for kw in ["参数组标识", "组标识", "切换参数组"]):
                    unique_vals = sorted(set(
                        str(v) for v in df[f].unique()
                        if pd.notna(v) and str(v).strip() != ''
                    ))
                    current_group_vals = self.group_filters.get(fid, {}).get(cmd, {}).get(f, set())

                    for val in unique_vals:
                        is_val_checked = val in current_group_vals
                        v_prefix = "[v] " if is_val_checked else "[ ] "
                        val_iid = f"{cmd}|||{f}|||{val}"
                        self.param_tree.insert(parent_node, "end", iid=val_iid,
                                              text=f"    {v_prefix}{str(val)}")

    def _on_tree_double_click(self, event):
        """双击切换勾选状态"""
        sel_file = self.file_lb.curselection()
        if not sel_file:
            return

        idx = sel_file[0]
        fid = self.files_data[idx]['id']
        table_dict = self.files_data[idx]['table_dict']

        item_id = self.param_tree.identify_row(event.y)
        if not item_id:
            return

        element = self.param_tree.identify_element(event.x, event.y)
        if 'indicator' in str(element):
            return

        parts = item_id.split("|||")

        if len(parts) == 3:
            cmd, field, val = parts
            self.group_filters.setdefault(fid, {}).setdefault(cmd, {}).setdefault(field, set())

            if val in self.group_filters[fid][cmd][field]:
                self.group_filters[fid][cmd][field].remove(val)
            else:
                self.group_filters[fid][cmd][field].add(val)
                self.selected_fields[fid][cmd].add(field)

            self.show_fields_tree(idx)
            return

        if len(parts) == 2:
            cmd, field = parts
            if field in self.selected_fields[fid][cmd]:
                self.selected_fields[fid][cmd].remove(field)
            else:
                self.selected_fields[fid][cmd].add(field)
            self.show_fields_tree(idx)
            return

        cmd = item_id
        all_fields = [c for c in table_dict[cmd].columns if c not in Constants.EXCLUDE_TREE_KEYS]
        if not all_fields:
            return

        if len(self.selected_fields[fid][cmd]) == len(all_fields):
            self.selected_fields[fid][cmd] = set()
        else:
            self.selected_fields[fid][cmd] = set(all_fields)
        self.show_fields_tree(idx)

    def select_all(self):
        sel = self.file_lb.curselection()
        if not sel:
            return
        idx = sel[0]
        fid = self.files_data[idx]['id']
        table_dict = self.files_data[idx]['table_dict']

        for cmd in table_dict.keys():
            all_fields = [c for c in table_dict[cmd].columns if c not in Constants.EXCLUDE_TREE_KEYS]
            self.selected_fields[fid][cmd] = set(all_fields)
        self.show_fields_tree(idx)

    def deselect_all(self):
        sel = self.file_lb.curselection()
        if not sel:
            return
        idx = sel[0]
        fid = self.files_data[idx]['id']

        if fid in self.selected_fields:
            for cmd in self.selected_fields[fid]:
                self.selected_fields[fid][cmd] = set()
        if fid in self.group_filters:
            self.group_filters[fid] = {}
        self.show_fields_tree(idx)

    def apply_config_to_all(self):
        sel = self.file_lb.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择母版文件。")
            return

        current_idx = sel[0]
        current_fid = self.files_data[current_idx]['id']
        template_cfg = self.selected_fields[current_fid]
        template_group_filters = self.group_filters.get(current_fid, {})

        for file_info in self.files_data:
            fid = file_info['id']
            if fid == current_fid:
                continue

            self.selected_fields[fid] = {}
            self.group_filters[fid] = {}

            t_dict = file_info['table_dict']
            for cmd in t_dict.keys():
                if cmd in template_cfg:
                    valid_cols = set(t_dict[cmd].columns)
                    self.selected_fields[fid][cmd] = template_cfg[cmd] & valid_cols
                    if cmd in template_group_filters:
                        for field, values in template_group_filters[cmd].items():
                            if field in valid_cols:
                                self.group_filters[fid].setdefault(cmd, {})[field] = values.copy()
                else:
                    self.selected_fields[fid][cmd] = set()

        messagebox.showinfo("同步完成", f"已成功将配置分发给其余 {len(self.files_data)-1} 个文件。")

    def save_config(self):
        sel = self.file_lb.curselection()
        if not sel:
            return
        idx = sel[0]
        fid = self.files_data[idx]['id']

        cfg_cols = {cmd: list(fields) for cmd, fields in self.selected_fields.get(fid, {}).items() if fields}
        cfg_groups = {}
        if fid in self.group_filters:
            for cmd, fields_dict in self.group_filters[fid].items():
                for field, values in fields_dict.items():
                    if values:
                        cfg_groups.setdefault(cmd, {})[field] = list(values)

        cfg = {"columns": cfg_cols, "group_filters": cfg_groups}

        if not cfg_cols and not cfg_groups:
            messagebox.showwarning("提示", "当前文件未配置任何有效的过滤规则")
            return

        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Config", "*.json")])
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "字段过滤配置文件导出完毕！")

    def load_config(self):
        """加载配置文件，增加JSON内容验证"""
        sel = self.file_lb.curselection()
        if not sel:
            return
        idx = sel[0]
        fid = self.files_data[idx]['id']
        table_dict = self.files_data[idx]['table_dict']

        path = filedialog.askopenfilename(filetypes=[("JSON Config", "*.json")])
        if not path:
            return

        # 安全校验路径
        if not _is_safe_path(path):
            messagebox.showerror("安全警告", f"不安全的配置文件路径被拒绝：{path}")
            return

        try:
            with open(path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)

            # JSON 结构验证
            if not isinstance(cfg, dict):
                raise ValueError("配置文件必须是 JSON 对象")
            
            if "columns" in cfg:
                cfg_cols = cfg["columns"]
                cfg_groups = cfg.get("group_filters", {})
                
                # 验证 columns 结构
                if not isinstance(cfg_cols, dict):
                    raise ValueError("'columns' 必须是字典类型")
                for cmd, fields in cfg_cols.items():
                    if not isinstance(fields, list):
                        raise ValueError(f"命令 '{cmd}' 的字段列表必须是数组类型")
                        
                # 验证 group_filters 结构
                if not isinstance(cfg_groups, dict):
                    raise ValueError("'group_filters' 必须是字典类型")
            else:
                cfg_cols = cfg
                cfg_groups = {}

            self.selected_fields[fid] = {}
            self.group_filters[fid] = {}

            for cmd in table_dict.keys():
                self.selected_fields[fid][cmd] = set()
                if cmd in cfg_cols:
                    valid_cols = set(table_dict[cmd].columns)
                    matched_fields = set(cfg_cols[cmd]) & valid_cols
                    self.selected_fields[fid][cmd] = matched_fields

                if cmd in cfg_groups:
                    for field, values in cfg_groups[cmd].items():
                        if field in table_dict[cmd].columns:
                            self.group_filters[fid].setdefault(cmd, {})[field] = set(str(v) for v in values)

            self.show_fields_tree(idx)
            messagebox.showinfo("配置导入成功", "成功匹配并导入配置文件属性与过滤项！")

        except json.JSONDecodeError as e:
            messagebox.showerror("加载失败", f"JSON 格式错误:\n{str(e)}")
        except ValueError as e:
            messagebox.showerror("加载失败", f"配置验证失败:\n{str(e)}")
        except FileNotFoundError:
            messagebox.showerror("加载失败", f"配置文件不存在:\n{path}")
        except PermissionError:
            messagebox.showerror("加载失败", f"无权限读取配置文件:\n{path}")
        except Exception as e:
            logger.error(f"加载配置异常：{e}")
            messagebox.showerror("加载失败", f"解析配置文件失败:\n{str(e)}")

        def export_all_merged(self):
            if not self.files_data:
                messagebox.showwarning("提示", "请先加载MML文件")
                return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return

        def worker():
            try:
                self.root.after(0, lambda: self.status.config(text="正在执行多维对齐合并..."))
                self.root.after(0, self.root.update_idletasks)

                all_file_dfs = []
                for file_info in self.files_data:
                    fid = file_info['id']
                    filtered_table_dict = apply_row_filters(file_info['table_dict'], self.group_filters, fid)

                    user_selection = self.selected_fields.get(fid, None)
                    if user_selection and not any(user_selection.values()):
                        user_selection = None

                    single_aligned_df = align_tables_by_cell(filtered_table_dict, user_selection)
                    if not single_aligned_df.empty:
                        all_file_dfs.append(single_aligned_df)

                if not all_file_dfs:
                    self.root.after(0, lambda: messagebox.showwarning("提示", "未提取到有效数据。"))
                    self.root.after(0, lambda: self.status.config(text="就绪"))
                    return

                combined_final_df = pd.concat(all_file_dfs, ignore_index=True)
                combined_final_df.drop_duplicates(inplace=True)
                success, msg = export_to_xlsx_format(combined_final_df, path)

                if success:
                    self.root.after(0, lambda: messagebox.showinfo("完成", msg))
                else:
                    self.root.after(0, lambda: messagebox.showerror("写入失败", msg))
                self.root.after(0, lambda: self.status.config(text="就绪"))

            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", err))
                self.root.after(0, lambda: self.status.config(text="就绪"))

        threading.Thread(target=worker, daemon=True).start()

    def match_5g_and_export(self):
        if not self.files_data:
            messagebox.showwarning("提示", "请先加载MML文件")
            return

        macro_filename = "5G参数处理宏.xlsm"
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            base_dir = os.getcwd()

        macro_path = os.path.join(base_dir, macro_filename)
        if not os.path.exists(macro_path):
            macro_path = os.path.join(os.getcwd(), macro_filename)

        if not os.path.exists(macro_path):
            messagebox.showerror("文件缺失",
                f"未能找到 {macro_filename}！\n请确保该文件与本程序放在同一文件夹下。")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return

        def worker():
            try:
                self.root.after(0, lambda: self.status.config(
                    text=f"正在读取 {macro_filename} 并执行匹配..."))
                self.root.after(0, self.root.update_idletasks)

                try:
                    df_macro = pd.read_excel(macro_path, sheet_name="基础信息", engine="openpyxl")
                except Exception as e:
                    self.root.after(0, lambda err=str(e): messagebox.showerror(
                        "读取宏文件失败", f"无法读取基础信息表:\n{err}"))
                    self.root.after(0, lambda: self.status.config(text="就绪"))
                    return

                all_file_dfs = []
                for file_info in self.files_data:
                    fid = file_info['id']
                    filtered_table_dict = apply_row_filters(
                        file_info['table_dict'], self.group_filters, fid)
                    user_selection = self.selected_fields.get(fid, None)
                    if user_selection and not any(user_selection.values()):
                        user_selection = None
                    single_aligned_df = align_tables_by_cell(filtered_table_dict, user_selection)
                    if not single_aligned_df.empty:
                        all_file_dfs.append(single_aligned_df)

                if not all_file_dfs:
                    self.root.after(0, lambda: messagebox.showwarning("提示", "未提取到有效数据。"))
                    self.root.after(0, lambda: self.status.config(text="就绪"))
                    return

                combined_final_df = pd.concat(all_file_dfs, ignore_index=True)
                combined_final_df.drop_duplicates(inplace=True)

                match_left_col = None
                if 'NR DU小区名称' in combined_final_df.columns:
                    match_left_col = 'NR DU小区名称'
                elif '小区名称' in combined_final_df.columns:
                    match_left_col = '小区名称'

                if match_left_col:
                    match_col = df_macro.columns[0]
                    target_cols = ["框号", "RRU型号", "反开4G", "是否扩容", "区分"]
                    actual_cols = [c for c in target_cols if c in df_macro.columns]

                    if actual_cols:
                        df_macro_unique = df_macro.drop_duplicates(subset=[match_col])
                        macro_subset = df_macro_unique[[match_col] + actual_cols]

                        combined_final_df = combined_final_df.merge(
                            macro_subset,
                            left_on=match_left_col,
                            right_on=match_col,
                            how='left'
                        )

                        if match_col != match_left_col:
                            combined_final_df.drop(columns=[match_col], inplace=True)

                success, msg = export_to_xlsx_format(combined_final_df, path)
                if success:
                    self.root.after(0, lambda: messagebox.showinfo("完成", msg))
                else:
                    self.root.after(0, lambda: messagebox.showerror("写入失败", msg))
                self.root.after(0, lambda: self.status.config(text="就绪"))

            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", err))
                self.root.after(0, lambda: self.status.config(text="就绪"))

        threading.Thread(target=worker, daemon=True).start()

    def export_all_per_sheet(self):
        if not self.files_data:
            messagebox.showwarning("提示", "请先加载MML文件")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return

        def worker():
            try:
                self.root.after(0, lambda: self.status.config(text="正在执行分Sheet导出..."))
                self.root.after(0, self.root.update_idletasks)

                cmd_matrix_groups = {}
                for file_info in self.files_data:
                    fid = file_info['id']
                    filtered_table_dict = apply_row_filters(
                        file_info['table_dict'], self.group_filters, fid)

                    for cmd_type, df in filtered_table_dict.items():
                        if not df.empty:
                            cmd_matrix_groups.setdefault(cmd_type, []).append(df)

                if not cmd_matrix_groups:
                    self.root.after(0, lambda: messagebox.showwarning("提示", "未解析到任何有效数据。"))
                    self.root.after(0, lambda: self.status.config(text="就绪"))
                    return

                final_sheets_dict = {}
                for cmd_type, dfs_list in cmd_matrix_groups.items():
                    combined_cmd_df = pd.concat(dfs_list, ignore_index=True)
                    combined_cmd_df.drop_duplicates(inplace=True)
                    final_sheets_dict[cmd_type] = combined_cmd_df

                success, msg = export_to_multi_sheets(final_sheets_dict, path)
                if success:
                    self.root.after(0, lambda: messagebox.showinfo("完成", msg))
                else:
                    self.root.after(0, lambda: messagebox.showerror("写入失败", msg))
                self.root.after(0, lambda: self.status.config(text="就绪"))

            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", err))
                self.root.after(0, lambda: self.status.config(text="就绪"))

        threading.Thread(target=worker, daemon=True).start()

# ==========================================
# 8. 主入口
# ==========================================
if __name__ == "__main__":

    root = tk.Tk()
    app = MMLToolGUI(root)
    root.mainloop()