#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 数据分类工具

按第一个Excel文件的指定类别列，对第二个Excel文件的指定列进行匹配分类，
每个类别输出到一个sheet，提取整行数据。支持精确匹配、包含匹配和模糊匹配。

用法:
    图形界面（默认）:  python fcp.py
    调试模式:          python fcp.py --debug
    命令行模式:        python fcp.py --cli 11.xlsx 22.xlsx --col1 A --col2 B -o out.xlsx
    命令行+调试:       python fcp.py --cli --debug 11.xlsx 22.xlsx --col1 A --col2 B

依赖:
    pip install -r requirements.txt
"""

import argparse
import re
import sys
import os
import threading
import traceback
from pathlib import Path

import pandas as pd

try:
    from rapidfuzz import fuzz, process
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False
    import difflib


# =============================================================================
# 核心函数
# =============================================================================

def normalize(s):
    """将值转换为用于匹配的标准化字符串。"""
    if s is None:
        return ""
    try:
        if pd.isna(s):
            return ""
    except (TypeError, ValueError):
        pass
    return str(s).strip().lower()


def sanitize_sheet_name(name, used_names):
    """生成合法且不重复的Excel sheet名称。"""
    name = str(name).strip()
    name = re.sub(r'[\\/*?:\[\]]', '_', name)
    name = name.strip("'")
    if not name:
        name = "未命名"
    name = name[:31]

    if name in used_names:
        i = 2
        # 预留 "_NNN" 后缀空间（4字符），确保总长度不超过31
        # 这样即使有999个重名也不会超限
        base = name[:27]
        candidate = f"{base}_{i}"
        while candidate in used_names:
            i += 1
            candidate = f"{base}_{i}"
        name = candidate

    used_names.add(name)
    return name


def parse_column(col_str):
    """
    解析Excel列号，支持 A, B, C, ..., Z, AA, AB, ... 格式。
    返回 0-based 列索引。
    """
    if not col_str:
        raise ValueError("列号不能为空")

    col_str = col_str.strip().upper()

    if not re.match(r'^[A-Z]+$', col_str):
        raise ValueError(f"无效的列号格式: '{col_str}'，应为 A, B, C, ..., Z, AA, AB 等")

    result = 0
    for char in col_str:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1


def parse_row_range(range_str, total_rows):
    """
    解析行范围参数，格式为 "起始行,终止行"（Excel行号，从1开始）。
    支持部分范围：
      - "2,100"  -> 第2行到第100行
      - "2,"     -> 第2行到末尾
      - ",100"   -> 第2行（数据起始）到第100行
    
    pandas read_excel 默认将第1行作为 header 消费，所以：
      Excel 第2行 = df.iloc[0]
      Excel 第3行 = df.iloc[1]
      转换公式: iloc = excel_row - 2
    
    返回 (start, end) 用于 pandas iloc 切片（0-based，end exclusive）。
    """
    if not range_str:
        return 0, total_rows
    
    parts = range_str.split(",")
    if len(parts) != 2:
        raise ValueError(f"行范围格式错误: '{range_str}'，应为 '起始行,终止行'，例如 '2,100'")
    
    start_str = parts[0].strip()
    end_str = parts[1].strip()
    
    # 解析起始行
    if start_str:
        try:
            start_excel = int(start_str)
        except ValueError:
            raise ValueError(f"起始行必须为整数: '{start_str}'")
        if start_excel < 2:
            raise ValueError(f"起始行不能小于2（第1行为表头），当前值: {start_excel}")
    else:
        start_excel = 2  # 默认从第2行（数据起始）
    
    # 解析终止行
    if end_str:
        try:
            end_excel = int(end_str)
        except ValueError:
            raise ValueError(f"终止行必须为整数: '{end_str}'")
        if end_excel < start_excel:
            raise ValueError(f"终止行({end_excel})不能小于起始行({start_excel})")
    else:
        end_excel = None  # 表示到末尾
    
    # 统一转换公式：Excel 行号 -> iloc 索引
    start_iloc = start_excel - 2
    if end_excel is not None:
        end_iloc = end_excel - 1  # 包含终止行，iloc exclusive 所以 -1
    else:
        end_iloc = total_rows  # 到末尾
    
    # 截断到实际数据范围
    start_iloc = min(start_iloc, total_rows)
    end_iloc = min(end_iloc, total_rows)
    return start_iloc, end_iloc


def find_best_match(value, categories, cat_norms, exact_map, threshold, debug=False):
    """
    返回(matched_category, score)。
    匹配顺序：精确匹配 -> 包含匹配 -> 模糊匹配。
    
    参数:
        value: 待匹配的值
        categories: 原始类别列表
        cat_norms: 预计算好的归一化类别列表（与 categories 一一对应）
        exact_map: 精确匹配字典 {normalized: original}，O(1) 查找
        threshold: 模糊匹配阈值
        debug: 是否打印调试日志
    """
    val_norm = normalize(value)
    if not val_norm:
        return None, 0

    # 1) 精确匹配 - O(1) 字典查找
    if val_norm in exact_map:
        cat = exact_map[val_norm]
        if debug:
            print(f"  [DEBUG] 精确匹配: '{value}' -> '{cat}'")
        return cat, 100

    # 2) 包含匹配
    best_cat = None
    best_len = -1
    for cat, cat_norm in zip(categories, cat_norms):
        if not cat_norm:
            continue
        if cat_norm in val_norm or val_norm in cat_norm:
            if len(cat_norm) > best_len:
                best_len = len(cat_norm)
                best_cat = cat
    if best_cat is not None:
        if debug:
            print(f"  [DEBUG] 包含匹配: '{value}' -> '{best_cat}'")
        return best_cat, 100

    # 3) 模糊匹配
    if HAS_RAPIDFUZZ:
        result = process.extractOne(val_norm, cat_norms, scorer=fuzz.WRatio)
        if result:
            choice, score, idx = result
            if score >= threshold:
                if debug:
                    print(f"  [DEBUG] 模糊匹配: '{value}' -> '{categories[idx]}' (score={score:.1f})")
                return categories[idx], score
    else:
        best_cat = None
        best_score = 0.0
        for cat, cat_norm in zip(categories, cat_norms):
            score = difflib.SequenceMatcher(None, val_norm, cat_norm).ratio() * 100
            if score > best_score:
                best_score = score
                best_cat = cat
        if best_score >= threshold:
            if debug:
                print(f"  [DEBUG] 模糊匹配(difflib): '{value}' -> '{best_cat}' (score={best_score:.1f})")
            return best_cat, best_score

    if debug:
        print(f"  [DEBUG] 未匹配: '{value}'")
    return None, 0


class ClassificationCancelled(Exception):
    """用户取消分类任务时抛出。"""
    pass


def classify_excel(file1, file2, col1="A", col2="B", rows1=None, rows2=None,
                   output="classified.xlsx", threshold=70, skip_empty=False,
                   unmatched_sheet="未匹配", progress_callback=None,
                   cancel_check=None, debug=False):
    """
    核心分类函数。

    参数:
        file1: 第一个Excel文件路径（类别文件）
        file2: 第二个Excel文件路径（数据文件）
        col1: 文件1中类别列的列号（Excel格式，如 A, B, C），默认 "A"
        col2: 文件2中待匹配列的列号（Excel格式，如 A, B, C），默认 "B"
        rows1: 文件1行范围，格式 "起始行,终止行"（Excel行号）或 None
        rows2: 文件2行范围，格式 "起始行,终止行"（Excel行号）或 None
        output: 输出文件路径
        threshold: 模糊匹配阈值（0-100）
        skip_empty: 是否跳过空sheet
        unmatched_sheet: 未匹配sheet名称
        progress_callback: 进度回调函数，签名 callback(message, percent)
        cancel_check: 取消检查函数，返回 True 表示用户已取消
        debug: 是否在控制台打印调试日志

    返回:
        dict: 包含统计信息的字典

    异常:
        ClassificationCancelled: 用户取消任务时抛出
    """
    def log(msg, percent=None):
        if debug:
            print(f"[DEBUG] {msg}")
        if progress_callback:
            progress_callback(msg, percent)
        elif not debug:
            print(msg)

    if not HAS_RAPIDFUZZ:
        log("提示: 未安装 rapidfuzz，将使用 difflib 进行模糊匹配。")

    # 解析列号
    try:
        col1_idx = parse_column(col1)
        col2_idx = parse_column(col2)
    except ValueError as e:
        raise ValueError(f"列号解析失败: {e}")

    if debug:
        print(f"[DEBUG] 列号解析: col1='{col1}'->idx {col1_idx}, col2='{col2}'->idx {col2_idx}")

    # 读取Excel
    log("正在读取Excel文件...", 5)
    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
    except Exception as e:
        raise ValueError(f"读取Excel文件失败: {e}")

    if debug:
        print(f"[DEBUG] 文件1: {len(df1)} 行, {len(df1.columns)} 列, 列名={list(df1.columns)}")
        print(f"[DEBUG] 文件2: {len(df2)} 行, {len(df2.columns)} 列, 列名={list(df2.columns)}")

    # 验证列号
    if col1_idx >= len(df1.columns):
        raise ValueError(f"文件1列号 '{col1}' 超出范围，该文件只有 {len(df1.columns)} 列")
    if col2_idx >= len(df2.columns):
        raise ValueError(f"文件2列号 '{col2}' 超出范围，该文件只有 {len(df2.columns)} 列")

    # 应用行范围
    try:
        if rows1:
            s1, e1 = parse_row_range(rows1, len(df1))
            df1 = df1.iloc[s1:e1].reset_index(drop=True)
            log(f"文件1 应用行范围 {rows1}，取到 {len(df1)} 行", 15)
        if rows2:
            s2, e2 = parse_row_range(rows2, len(df2))
            df2 = df2.iloc[s2:e2].reset_index(drop=True)
            log(f"文件2 应用行范围 {rows2}，取到 {len(df2)} 行", 20)
    except ValueError as e:
        raise ValueError(str(e))

    # 获取列名
    col1_name = df1.columns[col1_idx]
    col2_name = df2.columns[col2_idx]

    if debug:
        print(f"[DEBUG] 文件1类别列: col={col1}, name='{col1_name}'")
        print(f"[DEBUG] 文件2匹配列: col={col2}, name='{col2_name}'")

    # 第一个文件的类别列：去空、去重、保留顺序
    categories = df1[col1_name].dropna().astype(str).str.strip().unique().tolist()
    categories = [c for c in categories if c]
    if not categories:
        raise ValueError("第一个文件的分类列为空，请检查文件内容。")

    log(f"读取到 {len(categories)} 个类别", 25)
    log(f"文件1 类别列: {col1} ({col1_name})", 27)
    log(f"文件2 匹配列: {col2} ({col2_name})", 28)
    log(f"待分类数据共 {len(df2)} 行", 30)

    if debug:
        print(f"[DEBUG] 类别列表: {categories}")

    # 预计算归一化类别，避免每行重复 normalize
    cat_norms = [normalize(c) for c in categories]
    
    # 精确匹配优化：O(1) 字典查找替代 O(n) 线性扫描
    exact_map = {}
    for cat, cat_norm in zip(categories, cat_norms):
        if cat_norm and cat_norm not in exact_map:
            exact_map[cat_norm] = cat

    matched = {cat: [] for cat in categories}
    unmatched_rows = []
    cache = {}
    total_rows = len(df2)

    # 性能优化：使用 itertuples 替代 iterrows（快100倍）
    cols = df2.columns.tolist()
    col2_pos = cols.index(col2_name)
    
    # 逐行匹配
    log("开始匹配...", 35)
    for idx, row_tuple in enumerate(df2.itertuples(index=False)):
        # 检查用户是否取消
        if cancel_check and cancel_check():
            log("用户已取消任务", 0)
            raise ClassificationCancelled("用户取消")

        val = row_tuple[col2_pos]
        key = normalize(val)

        if key not in cache:
            cat, score = find_best_match(val, categories, cat_norms, exact_map, threshold, debug=debug)
            cache[key] = (cat, score)
        else:
            cat, score = cache[key]

        row_dict = dict(zip(cols, row_tuple))

        if cat is not None:
            matched[cat].append(row_dict)
        else:
            unmatched_rows.append(row_dict)

        # 更新进度（35% - 85%）
        if total_rows > 0 and ((idx + 1) % max(1, total_rows // 20) == 0 or idx == total_rows - 1):
            percent = 35 + int((idx + 1) / total_rows * 50)
            log(f"匹配进度: {idx + 1}/{total_rows}", percent)

    # 输出
    log("正在生成输出文件...", 90)
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    output_columns = df2.columns.tolist()

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            used_sheet_names = set()

            for cat in categories:
                rows = matched[cat]
                if skip_empty and not rows:
                    continue

                df_cat = pd.DataFrame(rows, columns=output_columns)
                sheet_name = sanitize_sheet_name(cat, used_sheet_names)
                df_cat.to_excel(writer, sheet_name=sheet_name, index=False)

            if unmatched_rows:
                sheet_name = sanitize_sheet_name(unmatched_sheet, used_sheet_names)
                pd.DataFrame(unmatched_rows, columns=output_columns).to_excel(
                    writer, sheet_name=sheet_name, index=False
                )
    except PermissionError:
        raise ValueError(f"无法写入输出文件：{output_path}\n文件可能正在被其他程序（如Excel）占用，请关闭后重试。")

    log(f"已完成，输出文件：{output_path}", 100)

    result = {
        "output": str(output_path),
        "categories": {cat: len(matched[cat]) for cat in categories},
        "unmatched": len(unmatched_rows),
        "total": total_rows
    }
    return result


# =============================================================================
# GUI
# =============================================================================

def run_gui(debug=False):
    """启动图形界面。"""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext

    class ExcelClassifierGUI:
        def __init__(self, root, debug=False):
            self.root = root
            self.debug = debug
            self.root.title("Excel 数据分类工具")
            self.root.geometry("750x720")
            self.root.resizable(True, True)

            # 变量
            self.file1_path = tk.StringVar()
            self.file2_path = tk.StringVar()
            self.col1 = tk.StringVar(value="A")
            self.col2 = tk.StringVar(value="B")
            self.rows1_start = tk.StringVar(value="2")
            self.rows1_end = tk.StringVar(value="")
            self.rows2_start = tk.StringVar(value="2")
            self.rows2_end = tk.StringVar(value="")
            self.output_path = tk.StringVar(value="classified.xlsx")
            self.threshold = tk.StringVar(value="70")
            self.is_running = False

            self.setup_ui()

            if self.debug:
                print("[DEBUG] GUI 初始化完成")

        def setup_ui(self):
            main_frame = ttk.Frame(self.root, padding="10")
            main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(0, weight=1)
            main_frame.columnconfigure(1, weight=1)

            row = 0

            # 文件1
            ttk.Label(main_frame, text="类别文件 (文件1):").grid(row=row, column=0, sticky=tk.W, pady=5)
            f1 = ttk.Frame(main_frame)
            f1.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
            f1.columnconfigure(0, weight=1)
            ttk.Entry(f1, textvariable=self.file1_path).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
            ttk.Button(f1, text="浏览...", command=self.browse_file1).grid(row=0, column=1)
            row += 1

            # 文件2
            ttk.Label(main_frame, text="数据文件 (文件2):").grid(row=row, column=0, sticky=tk.W, pady=5)
            f2 = ttk.Frame(main_frame)
            f2.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
            f2.columnconfigure(0, weight=1)
            ttk.Entry(f2, textvariable=self.file2_path).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
            ttk.Button(f2, text="浏览...", command=self.browse_file2).grid(row=0, column=1)
            row += 1

            ttk.Separator(main_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
            row += 1

            # 文件1 设置
            ttk.Label(main_frame, text="文件1 设置:").grid(row=row, column=0, sticky=tk.W, pady=5)
            c1 = ttk.Frame(main_frame)
            c1.grid(row=row, column=1, sticky=tk.W, pady=5)
            ttk.Label(c1, text="类别列号:").pack(side=tk.LEFT, padx=(0, 5))
            ttk.Entry(c1, textvariable=self.col1, width=5).pack(side=tk.LEFT, padx=(0, 15))
            ttk.Label(c1, text="起始行:").pack(side=tk.LEFT, padx=(0, 5))
            ttk.Entry(c1, textvariable=self.rows1_start, width=8).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Label(c1, text="终止行:").pack(side=tk.LEFT, padx=(0, 5))
            ttk.Entry(c1, textvariable=self.rows1_end, width=8).pack(side=tk.LEFT)
            ttk.Label(c1, text="(留空=全部)", foreground='gray').pack(side=tk.LEFT, padx=10)
            row += 1

            # 文件2 设置
            ttk.Label(main_frame, text="文件2 设置:").grid(row=row, column=0, sticky=tk.W, pady=5)
            c2 = ttk.Frame(main_frame)
            c2.grid(row=row, column=1, sticky=tk.W, pady=5)
            ttk.Label(c2, text="匹配列号:").pack(side=tk.LEFT, padx=(0, 5))
            ttk.Entry(c2, textvariable=self.col2, width=5).pack(side=tk.LEFT, padx=(0, 15))
            ttk.Label(c2, text="起始行:").pack(side=tk.LEFT, padx=(0, 5))
            ttk.Entry(c2, textvariable=self.rows2_start, width=8).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Label(c2, text="终止行:").pack(side=tk.LEFT, padx=(0, 5))
            ttk.Entry(c2, textvariable=self.rows2_end, width=8).pack(side=tk.LEFT)
            ttk.Label(c2, text="(留空=全部)", foreground='gray').pack(side=tk.LEFT, padx=10)
            row += 1

            # 提示
            hint = ttk.Frame(main_frame)
            hint.grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=2)
            ttk.Label(hint, text="提示: 列号为Excel列字母(A,B,C,...,AA,AB...)，行号为Excel行号(第1行为表头)",
                      foreground='gray').pack(side=tk.LEFT)
            row += 1

            ttk.Separator(main_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
            row += 1

            # 输出文件
            ttk.Label(main_frame, text="输出文件:").grid(row=row, column=0, sticky=tk.W, pady=5)
            of = ttk.Frame(main_frame)
            of.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
            of.columnconfigure(0, weight=1)
            ttk.Entry(of, textvariable=self.output_path).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
            ttk.Button(of, text="浏览...", command=self.browse_output).grid(row=0, column=1)
            row += 1

            # 阈值
            ttk.Label(main_frame, text="模糊匹配阈值:").grid(row=row, column=0, sticky=tk.W, pady=5)
            tf = ttk.Frame(main_frame)
            tf.grid(row=row, column=1, sticky=tk.W, pady=5)
            ttk.Entry(tf, textvariable=self.threshold, width=10).pack(side=tk.LEFT)
            ttk.Label(tf, text="(0-100，默认70，中文建议60-70)", foreground='gray').pack(side=tk.LEFT, padx=10)
            row += 1

            ttk.Separator(main_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
            row += 1

            # 进度条
            self.progress_var = tk.DoubleVar()
            self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
            self.progress_bar.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
            row += 1

            # 日志
            ttk.Label(main_frame, text="运行日志:").grid(row=row, column=0, sticky=tk.W, pady=5)
            row += 1
            self.log_text = scrolledtext.ScrolledText(main_frame, height=10, width=90)
            self.log_text.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
            main_frame.rowconfigure(row, weight=1)
            row += 1

            # 按钮
            bf = ttk.Frame(main_frame)
            bf.grid(row=row, column=0, columnspan=2, pady=10)
            self.start_button = ttk.Button(bf, text="开始分类", command=self.start_classification)
            self.start_button.pack(side=tk.LEFT, padx=5)
            self.stop_button = ttk.Button(bf, text="停止", command=self.stop_classification, state=tk.DISABLED)
            self.stop_button.pack(side=tk.LEFT, padx=5)
            ttk.Button(bf, text="清空日志", command=self.clear_log).pack(side=tk.LEFT, padx=5)

        # ---------- 文件浏览 ----------

        def browse_file1(self):
            f = filedialog.askopenfilename(title="选择类别文件",
                                           filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
            if f:
                self.file1_path.set(f)

        def browse_file2(self):
            f = filedialog.askopenfilename(title="选择数据文件",
                                           filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
            if f:
                self.file2_path.set(f)

        def browse_output(self):
            f = filedialog.asksaveasfilename(title="选择输出文件", defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if f:
                self.output_path.set(f)

        # ---------- 日志 ----------

        def log(self, message, percent=None):
            if percent is not None:
                self.progress_var.set(percent)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)

        def clear_log(self):
            self.log_text.delete(1.0, tk.END)
            self.progress_var.set(0)

        # ---------- 验证 ----------

        def validate_inputs(self):
            if not self.file1_path.get():
                messagebox.showerror("错误", "请选择类别文件 (文件1)")
                return False
            if not self.file2_path.get():
                messagebox.showerror("错误", "请选择数据文件 (文件2)")
                return False
            if not os.path.exists(self.file1_path.get()):
                messagebox.showerror("错误", f"文件1不存在: {self.file1_path.get()}")
                return False
            if not os.path.exists(self.file2_path.get()):
                messagebox.showerror("错误", f"文件2不存在: {self.file2_path.get()}")
                return False
            if not self.output_path.get():
                messagebox.showerror("错误", "请指定输出文件路径")
                return False
            try:
                parse_column(self.col1.get())
            except ValueError as e:
                messagebox.showerror("错误", f"文件1列号格式错误: {e}")
                return False
            try:
                parse_column(self.col2.get())
            except ValueError as e:
                messagebox.showerror("错误", f"文件2列号格式错误: {e}")
                return False
            try:
                t = float(self.threshold.get())
                if not 0 <= t <= 100:
                    messagebox.showerror("错误", "模糊匹配阈值必须在 0-100 之间")
                    return False
            except ValueError:
                messagebox.showerror("错误", "模糊匹配阈值必须是数字")
                return False
            return True

        # ---------- 开始 / 停止 ----------

        def start_classification(self):
            if self.debug:
                print("[DEBUG] start_classification 被调用")

            if self.is_running:
                if self.debug:
                    print("[DEBUG] 已经在运行中，忽略")
                return

            if not self.validate_inputs():
                if self.debug:
                    print("[DEBUG] 输入验证失败")
                return

            # 先更新状态，再强制刷新 UI，最后才启动线程
            self.is_running = True
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.clear_log()
            self.root.update_idletasks()   # <-- 关键：强制渲染按钮状态

            if self.debug:
                print("[DEBUG] 按钮状态已更新，准备启动线程")

            t = threading.Thread(target=self._run_classification, daemon=True)
            t.start()

            if self.debug:
                print("[DEBUG] 线程已启动")

        def stop_classification(self):
            if self.debug:
                print("[DEBUG] stop_classification 被调用")
            self.is_running = False
            self.log("用户请求停止...")

        # ---------- 后台任务 ----------

        def _run_classification(self):
            try:
                if self.debug:
                    print("[DEBUG] _run_classification 开始执行")

                rows1 = None
                if self.rows1_start.get() or self.rows1_end.get():
                    rows1 = f"{self.rows1_start.get()},{self.rows1_end.get()}"

                rows2 = None
                if self.rows2_start.get() or self.rows2_end.get():
                    rows2 = f"{self.rows2_start.get()},{self.rows2_end.get()}"

                if self.debug:
                    print(f"[DEBUG] 参数: file1={self.file1_path.get()}, file2={self.file2_path.get()}")
                    print(f"[DEBUG] 参数: col1={self.col1.get()}, col2={self.col2.get()}")
                    print(f"[DEBUG] 参数: rows1={rows1}, rows2={rows2}")
                    print(f"[DEBUG] 参数: output={self.output_path.get()}, threshold={self.threshold.get()}")

                def progress_callback(message, percent):
                    # 始终在主线程中更新 UI
                    self.root.after(0, self._safe_log, message, percent)

                result = classify_excel(
                    file1=self.file1_path.get(),
                    file2=self.file2_path.get(),
                    col1=self.col1.get(),
                    col2=self.col2.get(),
                    rows1=rows1,
                    rows2=rows2,
                    output=self.output_path.get(),
                    threshold=float(self.threshold.get()),
                    progress_callback=progress_callback,
                    cancel_check=lambda: not self.is_running,
                    debug=self.debug,
                )

                if self.debug:
                    print(f"[DEBUG] classify_excel 返回: {result}")

                if self.is_running:
                    self.root.after(0, self._show_result, result)

            except ClassificationCancelled:
                if self.debug:
                    print("[DEBUG] 任务已被用户取消")
                # 取消不弹错误框，只记日志
                self.root.after(0, self._safe_log, "任务已取消", None)

            except Exception as e:
                if self.debug:
                    print(f"[DEBUG] 异常: {e}")
                    traceback.print_exc()
                if self.is_running:
                    self.root.after(0, self._show_error, str(e))

            finally:
                if self.debug:
                    print("[DEBUG] _run_classification 结束，重置 UI")
                self.root.after(0, self._reset_ui)

        # ---------- 主线程 UI 操作 ----------

        def _safe_log(self, message, percent):
            """在主线程中安全地写入日志。"""
            try:
                self.log(message, percent)
                self.root.update_idletasks()
            except Exception:
                pass

        def _show_result(self, result):
            self.log("\n" + "=" * 50)
            self.log("分类完成!")
            self.log("=" * 50)
            self.log(f"输出文件: {result['output']}")
            self.log(f"总数据行: {result['total']}")
            self.log("\n各类别统计:")
            for cat, count in result['categories'].items():
                if count > 0:
                    self.log(f"  {cat}: {count} 行")
            self.log(f"  未匹配: {result['unmatched']} 行")
            messagebox.showinfo("完成",
                                f"分类完成!\n\n输出文件: {result['output']}\n总数据行: {result['total']}")

        def _show_error(self, error_msg):
            self.log(f"\n错误: {error_msg}")
            messagebox.showerror("错误", f"分类失败:\n{error_msg}")

        def _reset_ui(self):
            self.is_running = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.root.update_idletasks()

    root = tk.Tk()
    app = ExcelClassifierGUI(root, debug=debug)
    root.mainloop()


# =============================================================================
# CLI
# =============================================================================

def run_cli(args, debug=False):
    """命令行模式。"""
    try:
        result = classify_excel(
            file1=args.file1,
            file2=args.file2,
            col1=args.col1,
            col2=args.col2,
            rows1=args.rows1,
            rows2=args.rows2,
            output=args.output,
            threshold=args.threshold,
            skip_empty=args.skip_empty,
            unmatched_sheet=args.unmatched_sheet,
            debug=debug,
        )

        print(f"\n统计结果：")
        for cat, count in result["categories"].items():
            print(f"  {cat}: {count} 行")
        print(f"  未匹配: {result['unmatched']} 行")
    except ValueError as e:
        sys.exit(str(e))


# =============================================================================
# 入口
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Excel 数据分类工具：按类别列对数据进行匹配分类，每个类别输出到一个sheet。",
        epilog="不带 --cli 时默认启动图形界面。"
    )
    parser.add_argument("--cli", action="store_true",
                        help="使用命令行模式（默认启动图形界面）")
    parser.add_argument("--debug", action="store_true",
                        help="调试模式：在控制台打印详细调试日志")

    # 以下参数仅在 --cli 模式下使用
    parser.add_argument("file1", nargs="?", help="第一个Excel文件（类别文件，CLI模式必填）")
    parser.add_argument("file2", nargs="?", help="第二个Excel文件（数据文件，CLI模式必填）")
    parser.add_argument("--col1", default="A",
                        help="文件1中类别列的列号（Excel格式，如 A, B, C），默认 A")
    parser.add_argument("--col2", default="B",
                        help="文件2中待匹配列的列号（Excel格式，如 A, B, C），默认 B")
    parser.add_argument("--rows1", default=None,
                        help="文件1的行范围，格式: 起始行,终止行（Excel行号），例如 2,100")
    parser.add_argument("--rows2", default=None,
                        help="文件2的行范围，格式: 起始行,终止行（Excel行号），例如 2,500")
    parser.add_argument("-o", "--output", default="classified.xlsx",
                        help="输出Excel文件路径，默认 classified.xlsx")
    parser.add_argument("-t", "--threshold", type=float, default=70,
                        help="模糊匹配阈值（0-100），默认70")
    parser.add_argument("--skip-empty", action="store_true",
                        help="跳过没有匹配行的类别sheet")
    parser.add_argument("--unmatched-sheet", default="未匹配",
                        help="未匹配行的sheet名称，默认：未匹配")

    args = parser.parse_args()

    if args.cli:
        if not args.file1 or not args.file2:
            parser.error("命令行模式需要指定 file1 和 file2")
        run_cli(args, debug=args.debug)
    else:
        run_gui(debug=args.debug)


if __name__ == "__main__":
    # EXE 打包适配：顶层异常捕获，防止 GUI 模式崩溃无提示
    try:
        main()
    except Exception as e:
        # 尝试显示 GUI 错误对话框
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("程序错误", f"程序发生未预期的错误：\n\n{type(e).__name__}: {e}\n\n详细信息已保存到 error.log")
            root.destroy()
        except:
            pass
        
        # 保存错误日志
        import traceback
        with open("error.log", "w", encoding="utf-8") as f:
            f.write(f"错误时间: {__import__('datetime').datetime.now()}\n")
            f.write(f"错误类型: {type(e).__name__}\n")
            f.write(f"错误信息: {e}\n\n")
            f.write("详细堆栈:\n")
            traceback.print_exc(file=f)
        
        # 如果是 CLI 模式，打印到控制台
        if "--cli" in sys.argv:
            traceback.print_exc()
        
        sys.exit(1)
