#!/usr/bin/env python3
"""
新华医院物业管理费及维修费分析（累积月份） — 从本年累计及上年同期累计物业维修明细 Excel 更新分析底稿。

本脚本处理的是累积数据（1月至当前月份的所有交易明细），与单月版本的区别：
- 输入文件包含的是 1月到当前月份的累计交易数据，而非仅当月数据
- 汇总后写入底稿的数据反映的是本年累计 vs 上年同期累计的对比

用法:
    python property_analysis.py <本年累计明细文件> <上年同期累计明细文件> <分析底稿> [--output OUTPUT] [--mode MODE]

    --mode property    : 仅更新物业管理费
    --mode maintenance : 仅更新维修费
    --mode all         : 同时更新物业管理费和维修费（默认）

示例:
    python property_analysis.py 2026年1-6月物业维修费.xlsx 2025年1-6月物业维修费.xls 6月分析底稿.xlsx
"""

import json
import os
import sys
import re
import argparse
from datetime import datetime
from collections import defaultdict

import openpyxl


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def norm(s):
    """标准化文本用于模糊匹配."""
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r'[\s　\n\r\t（）()]+', '', s)
    return s


def fuzzy_find(text, candidates):
    """在 candidates 列表中查找包含 text 或 被 text 包含的项."""
    t = norm(text)
    for c in candidates:
        cn = norm(c)
        if t and cn and (t in cn or cn in t):
            return c
    return None


# ---------------------------------------------------------------------------
# Cross-format cell access (openpyxl + xlrd)
# ---------------------------------------------------------------------------

def get_cell_value(ws, row_1based, col_1based, is_xlrd=False):
    """统一的单元格读取 — 兼容 openpyxl 和 xlrd."""
    if is_xlrd:
        # xlrd: 0-based row/col
        return ws.cell_value(row_1based - 1, col_1based - 1)
    else:
        # openpyxl: 1-based row/col
        return ws.cell(row_1based, col_1based).value


def load_workbook_any(path):
    """加载 xlsx (openpyxl) 或 xls (xlrd) 文件.
    Returns: (workbook, is_xlrd)
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == '.xls' and not path.lower().endswith('.xlsx'):
        import xlrd
        return xlrd.open_workbook(path), True
    else:
        return openpyxl.load_workbook(path, data_only=True), False


def get_sheet_by_index(wb, idx, is_xlrd=False):
    """按索引获取 worksheet."""
    if is_xlrd:
        return wb.sheet_by_index(idx)
    else:
        return wb.worksheets[idx]


# ---------------------------------------------------------------------------
# Format Detection & Data Loading
# ---------------------------------------------------------------------------

def detect_format(wb, is_xlrd=False):
    """检测物业明细文件的格式（2026新格式 vs 2025旧格式）."""
    ws = get_sheet_by_index(wb, 0, is_xlrd)

    h0 = str(get_cell_value(ws, 1, 1, is_xlrd) or '')
    h0_r3 = str(get_cell_value(ws, 3, 1, is_xlrd) or '')

    # 2026 format: row 3 col 1 = "年"
    r3c2 = str(get_cell_value(ws, 3, 2, is_xlrd) or '')
    if '年' in h0_r3 and '月' in r3c2:
        return {
            "category_col": 6,    # 0-based
            "campus_col": 7,
            "amount_col": 10,
            "code_col": 3,
            "name_col": 4,
            "data_start_row": 5,
            "has_prefix": False,
            "campus_prefix": "",
            "campus_map": {"杨浦院区": "杨浦", "奉贤院区": "奉贤"},
        }

    # 2025 format: row 2 col 1 = "科目代码"
    r2c1 = str(get_cell_value(ws, 2, 1, is_xlrd) or '')
    if '科目代码' in r2c1:
        return {
            "category_col": 11,
            "campus_col": 8,
            "amount_col": 12,
            "code_col": 0,
            "name_col": 1,
            "data_start_row": 3,
            "has_prefix": True,
            "category_prefix": "运营成本:物业管理_",
            "maintenance_category_prefix": "运营成本:维修维护_",
            "campus_prefix": "分院:",
            "campus_map": {"杨浦院区": "杨浦", "奉贤院区": "奉贤"},
        }

    # Fallback: try to detect by max_column
    if is_xlrd:
        max_col = ws.ncols
    else:
        max_col = ws.max_column

    if max_col >= 15:
        return {
            "category_col": 11,
            "campus_col": 8,
            "amount_col": 12,
            "code_col": 0,
            "name_col": 1,
            "data_start_row": 3,
            "has_prefix": True,
            "category_prefix": "运营成本:物业管理_",
            "maintenance_category_prefix": "运营成本:维修维护_",
            "campus_prefix": "分院:",
            "campus_map": {"杨浦院区": "杨浦", "奉贤院区": "奉贤"},
        }
    else:
        return {
            "category_col": 6,
            "campus_col": 7,
            "amount_col": 10,
            "code_col": 3,
            "name_col": 4,
            "data_start_row": 5,
            "has_prefix": False,
            "campus_prefix": "",
            "campus_map": {"杨浦院区": "杨浦", "奉贤院区": "奉贤"},
        }


def get_max_row(ws, is_xlrd=False):
    """获取 worksheet 的最大行数."""
    if is_xlrd:
        return ws.nrows
    else:
        return ws.max_row


def load_categories(ws, fmt, category_map, exclude_keywords, is_xlrd=False):
    """从明细 sheet 中按运营成本分类汇总金额（元）— 物业管理费.

    Returns:
        {category_label: {"杨浦": total, "奉贤": total}}
    """
    result = defaultdict(lambda: defaultdict(float))
    cat_col = fmt["category_col"]
    campus_col = fmt["campus_col"]
    amt_col = fmt["amount_col"]
    start_row = fmt["data_start_row"]
    has_prefix = fmt.get("has_prefix", False)
    category_prefix = fmt.get("category_prefix", "")
    campus_prefix = fmt.get("campus_prefix", "")
    campus_map = fmt.get("campus_map", {})
    max_row = get_max_row(ws, is_xlrd)

    for r in range(start_row, max_row + 1):
        raw_cat = str(get_cell_value(ws, r, cat_col + 1, is_xlrd) or '').strip()
        if not raw_cat:
            continue

        # Skip non-物业管理 rows
        if has_prefix:
            # 2025 format: must start with "运营成本:物业管理_"
            if not raw_cat.startswith(category_prefix):
                continue
            base_cat = raw_cat[len(category_prefix):]
        else:
            # 2026 format: check against exclude list
            base_cat = raw_cat
            is_excluded = False
            for kw in exclude_keywords:
                if kw in base_cat:
                    is_excluded = True
                    break
            if base_cat == "其他":
                is_excluded = True
            if is_excluded:
                continue

        # Fuzzy match to known category
        matched = None
        for cat_def in category_map:
            for kw in cat_def["keywords"]:
                if norm(kw) in norm(base_cat) or norm(base_cat) in norm(kw):
                    matched = cat_def["label"]
                    break
            if matched:
                break

        if matched is None:
            continue

        # Extract campus
        raw_campus = str(get_cell_value(ws, r, campus_col + 1, is_xlrd) or '').strip()
        if campus_prefix:
            raw_campus = raw_campus.replace(campus_prefix, '')
        campus = campus_map.get(raw_campus, raw_campus)
        if not campus:
            campus = "杨浦"

        # Sum amount
        amt = get_cell_value(ws, r, amt_col + 1, is_xlrd)
        if isinstance(amt, (int, float)):
            result[matched][campus] += float(amt)

    return result


# ---------------------------------------------------------------------------
# Maintenance Fee Classification
# ---------------------------------------------------------------------------

def classify_top_level(name, code_no_dot):
    """按科目名称将维修费行分类到一级类别（维修护费 sheet）.

    Args:
        name: 科目名称（2026短名或2025长层级名）
        code_no_dot: 去除点号的科目编码

    Returns:
        分类标签 (str) 或 None
    """
    n = norm(name)
    # Prioritized matching: specific categories before catch-all "其他"
    if "设备日常维修" in n:
        return "设备维修"
    if "网络信息系统运行与维护费" in n or "网络信息系统" in n:
        return "网络维修"
    if "房屋日常维修" in n:
        return "房屋维修"
    # 后勤维修: 科目名称含"其他" 且 科目编码以04结尾
    if "其他" in n and code_no_dot.endswith("04"):
        return "后勤维修"
    # 2025 format: 科目名称含"维修（护）费_其他"
    if "维修（护）费_其他" in n or "维修(护)费_其他" in n:
        return "后勤维修"
    return None


def classify_logistics(base_cat):
    """按运营成本将后勤维修行分类到二级类别（后勤维修 sheet）.

    Args:
        base_cat: 去除前缀后的运营成本文本

    Returns:
        分类标签 (str)
    """
    n = norm(base_cat)
    # Ordered from most specific to most general
    mapping = [
        ("物业维修服务费", "物业维修"),
        ("电梯运行管理费", "电梯运行"),
        ("电站运行管理费", "电站运行"),
        ("空调运行管理费", "空调运行管理费"),
        ("给排水锅炉运行管理费", "给排水锅炉"),
        ("后勤智能化平台", "后勤智能化"),
        ("污水处理站运行管理费", "污水处理站"),
        ("污水管线养护费", "污水管线"),
        ("医用气体运行管理费", "医用气体"),
        ("通用设施设备管理", "通用设施"),
    ]
    for kw, label in mapping:
        if kw in n:
            return label
    # Fallback: treat as "其他"
    return "其他"


def load_maintenance_data(ws, fmt, is_xlrd=False):
    """从明细 sheet 中加载维修费数据，同时完成一级和二级分类.

    Returns:
        {
            "top_level": {label: {"杨浦": total, "奉贤": total}},
            "logistics_repair": {label: {"杨浦": total, "奉贤": total}}
        }
    """
    cat_col = fmt["category_col"]
    campus_col = fmt["campus_col"]
    amt_col = fmt["amount_col"]
    code_col = fmt.get("code_col", 3)
    name_col = fmt.get("name_col", 4)
    start_row = fmt["data_start_row"]
    has_prefix = fmt.get("has_prefix", False)
    campus_prefix = fmt.get("campus_prefix", "")
    campus_map = fmt.get("campus_map", {})
    code_marker = "0412"
    maintenance_cat_prefix = fmt.get("maintenance_category_prefix",
                                       "运营成本:维修维护_")
    max_row = get_max_row(ws, is_xlrd)

    result_top = defaultdict(lambda: defaultdict(float))
    result_logistics = defaultdict(lambda: defaultdict(float))

    for r in range(start_row, max_row + 1):
        code = str(get_cell_value(ws, r, code_col + 1, is_xlrd) or '').strip()
        code_no_dot = code.replace('.', '').replace('-', '')

        # Filter: only 维修费 rows (code contains "0412")
        if code_marker not in code_no_dot:
            continue

        # Get 科目名称
        name = str(get_cell_value(ws, r, name_col + 1, is_xlrd) or '').strip()
        if not name:
            continue

        # Get 运营成本
        raw_cat = str(get_cell_value(ws, r, cat_col + 1, is_xlrd) or '').strip()
        if not raw_cat:
            continue

        # Strip prefix for 2025 format
        if has_prefix and raw_cat.startswith(maintenance_cat_prefix):
            base_cat = raw_cat[len(maintenance_cat_prefix):]
        else:
            base_cat = raw_cat

        # Extract campus
        raw_campus = str(get_cell_value(ws, r, campus_col + 1, is_xlrd) or '').strip()
        if campus_prefix:
            raw_campus = raw_campus.replace(campus_prefix, '')
        campus = campus_map.get(raw_campus, raw_campus)
        if not campus:
            campus = "杨浦"

        # Amount
        amt = get_cell_value(ws, r, amt_col + 1, is_xlrd)
        if not isinstance(amt, (int, float)):
            continue

        # Top-level classification by 科目名称
        top_cat = classify_top_level(name, code_no_dot)
        if top_cat is None:
            continue

        result_top[top_cat][campus] += float(amt)

        # Sub-level classification for 后勤维修 by 运营成本
        if top_cat == "后勤维修":
            sub_cat = classify_logistics(base_cat)
            result_logistics[sub_cat][campus] += float(amt)

    return {"top_level": dict(result_top), "logistics_repair": dict(result_logistics)}


# ---------------------------------------------------------------------------
# Sheet Updater — Common
# ---------------------------------------------------------------------------

def find_row_by_label(ws, label, search_col, max_row=None):
    """在指定列中查找包含 label 文本的行号（1-indexed）.

    search_col: 0-based column index to search in.
    """
    if max_row is None:
        max_row = ws.max_row
    target = norm(label)
    for r in range(1, max_row + 1):
        cell_text = norm(ws.cell(r, search_col + 1).value)
        if target and target in cell_text:
            return r
    # Looser: check if cell text contains parts of label
    for r in range(1, max_row + 1):
        cell_text = norm(ws.cell(r, search_col + 1).value)
        if target and cell_text and cell_text in target:
            return r
    return None


def _round_wan(amount_yuan):
    """元转万元，四舍五入."""
    return round(amount_yuan / 10000.0)


# ---------------------------------------------------------------------------
# Sheet Updater — Property Management (existing)
# ---------------------------------------------------------------------------

def update_workbook(wb, data_current, data_prev, category_map, targets, output_path):
    """更新分析底稿中的物业管理 sheet 和商品服务 sheet."""
    property_sheet_name = targets["property_sheet"]
    service_sheet_name = targets["service_sheet"]
    service_total_row = targets["service_total_row"]
    service_total_cols = targets["service_total_cols"]
    left_cfg = targets["left_table"]
    right_cfg = targets["right_table"]

    # -- Get sheets --
    if property_sheet_name not in [ws.title for ws in wb.worksheets]:
        print(f"[ERROR] Sheet '{property_sheet_name}' not found")
        return 0, 0
    ws_prop = wb[property_sheet_name]

    if service_sheet_name not in [ws.title for ws in wb.worksheets]:
        print(f"[ERROR] Sheet '{service_sheet_name}' not found")
        return 0, 0
    ws_svc = wb[service_sheet_name]

    updated_cells = 0
    skipped = 0

    # -- Compute totals --
    total_current_yangpu = 0.0
    total_current_fengxian = 0.0
    total_prev_yangpu = 0.0
    total_prev_fengxian = 0.0

    for cat_def in category_map:
        label = cat_def["label"]
        dc = data_current.get(label, {})
        dp = data_prev.get(label, {})
        total_current_yangpu += dc.get("杨浦", 0)
        total_current_fengxian += dc.get("奉贤", 0)
        total_prev_yangpu += dp.get("杨浦", 0)
        total_prev_fengxian += dp.get("奉贤", 0)

    tc_yangpu_wan = _round_wan(total_current_yangpu)
    tc_fengxian_wan = _round_wan(total_current_fengxian)
    tc_total_wan = _round_wan(total_current_yangpu + total_current_fengxian)
    tp_yangpu_wan = _round_wan(total_prev_yangpu)
    tp_fengxian_wan = _round_wan(total_prev_fengxian)
    tp_total_wan = _round_wan(total_prev_yangpu + total_prev_fengxian)

    print(f"\n  本年物业管理费合计: {tc_total_wan} 万元 (杨浦 {tc_yangpu_wan}, 奉贤 {tc_fengxian_wan})")
    print(f"  上年物业管理费合计: {tp_total_wan} 万元 (杨浦 {tp_yangpu_wan}, 奉贤 {tp_fengxian_wan})")

    # -- Update 商品服务 sheet R7 (物业管理费 total row) --
    svc_r = service_total_row
    svc_c = service_total_cols["current"] + 1  # 1-indexed
    svc_d = service_total_cols["fengxian"] + 1
    svc_e = service_total_cols["prev"] + 1

    old_c7 = ws_svc.cell(svc_r, svc_c).value
    old_d7 = ws_svc.cell(svc_r, svc_d).value
    old_e7 = ws_svc.cell(svc_r, svc_e).value

    ws_svc.cell(svc_r, svc_c).value = tc_total_wan
    ws_svc.cell(svc_r, svc_d).value = tc_fengxian_wan
    ws_svc.cell(svc_r, svc_e).value = tp_total_wan
    print(f"  [UPD] 商品服务 R{svc_r}: C={old_c7}->{tc_total_wan}, D={old_d7}->{tc_fengxian_wan}, E={old_e7}->{tp_total_wan}")
    updated_cells += 3

    # -- Update 物业管理 sheet --
    left_search_col = 1   # B列 (0-based)
    right_search_col = 8  # I列 (0-based)

    for cat_def in category_map:
        label = cat_def["label"]
        dc = data_current.get(label, {})
        dp = data_prev.get(label, {})

        cur_yangpu = _round_wan(dc.get("杨浦", 0))
        cur_fengxian = _round_wan(dc.get("奉贤", 0))
        cur_total = cur_yangpu + cur_fengxian
        prev_yangpu = _round_wan(dp.get("杨浦", 0))
        prev_fengxian = _round_wan(dp.get("奉贤", 0))
        prev_total = prev_yangpu + prev_fengxian

        # --- Left table ---
        if "left_row" in cat_def:
            lr = cat_def["left_row"]
            found = find_row_by_label(ws_prop, label, left_search_col)
            if found:
                lr = found
            else:
                print(f"  [SKIP] Left table: '{label}' row not found")
                skipped += 1
                continue

            lc = left_cfg["data_cols"]["current"] + 1   # C列
            ld = left_cfg["data_cols"]["fengxian"] + 1   # D列
            le = left_cfg["data_cols"]["prev"] + 1       # E列

            old_c = ws_prop.cell(lr, lc).value
            old_d = ws_prop.cell(lr, ld).value
            old_e = ws_prop.cell(lr, le).value

            ws_prop.cell(lr, lc).value = cur_total
            ws_prop.cell(lr, ld).value = cur_fengxian
            ws_prop.cell(lr, le).value = prev_total
            print(f"  [UPD] 物业管理 R{lr}C: {old_c}->{cur_total} | R{lr}D: {old_d}->{cur_fengxian} | R{lr}E: {old_e}->{prev_total}  [{label}]")
            updated_cells += 3

        # --- Right table ---
        if "right_row" in cat_def:
            rr = cat_def["right_row"]
            found = find_row_by_label(ws_prop, label, right_search_col,
                                       max_row=right_cfg["remainder_row"])
            if found:
                rr = found
            else:
                print(f"  [SKIP] Right table: '{label}' row not found")
                skipped += 1
                continue

            rc = right_cfg["data_cols"]["current"] + 1   # J列
            rd = right_cfg["data_cols"]["fengxian"] + 1   # K列
            rl = right_cfg["data_cols"]["prev"] + 1       # L列

            old_j = ws_prop.cell(rr, rc).value
            old_k = ws_prop.cell(rr, rd).value
            old_l = ws_prop.cell(rr, rl).value

            ws_prop.cell(rr, rc).value = cur_total
            ws_prop.cell(rr, rd).value = cur_fengxian
            ws_prop.cell(rr, rl).value = prev_total
            print(f"  [UPD] 物业管理 R{rr}J: {old_j}->{cur_total} | R{rr}K: {old_k}->{cur_fengxian} | R{rr}L: {old_l}->{prev_total}  [{label}]")
            updated_cells += 3

    print(f"\n  [物业管理] Summary: {updated_cells} cells updated, {skipped} rows skipped")
    return updated_cells, skipped


# ---------------------------------------------------------------------------
# Sheet Updater — Maintenance Fee
# ---------------------------------------------------------------------------

def update_service_row_for_maintenance(wb, data_top_current, data_top_prev, mt_config):
    """更新商品服务 sheet R6 维修费汇总行."""
    svc_config = mt_config["service_sheet"]
    sheet_name = svc_config["sheet"]

    if sheet_name not in [ws.title for ws in wb.worksheets]:
        print(f"[ERROR] Sheet '{sheet_name}' not found for maintenance service row")
        return 0

    ws = wb[sheet_name]
    r = svc_config["total_row"]
    c_cur = svc_config["data_cols"]["current"] + 1   # C列 (1-indexed)
    c_fx = svc_config["data_cols"]["fengxian"] + 1    # D列
    c_prev = svc_config["data_cols"]["prev"] + 1      # E列

    # Compute totals from top-level data
    total_cur = sum(
        d.get("杨浦", 0) + d.get("奉贤", 0)
        for d in data_top_current.values()
    )
    total_cur_fx = sum(
        d.get("奉贤", 0) for d in data_top_current.values()
    )
    total_prev = sum(
        d.get("杨浦", 0) + d.get("奉贤", 0)
        for d in data_top_prev.values()
    )

    cur_wan = _round_wan(total_cur)
    fx_wan = _round_wan(total_cur_fx)
    prev_wan = _round_wan(total_prev)

    old_c = ws.cell(r, c_cur).value
    old_d = ws.cell(r, c_fx).value
    old_e = ws.cell(r, c_prev).value

    ws.cell(r, c_cur).value = cur_wan
    ws.cell(r, c_fx).value = fx_wan
    ws.cell(r, c_prev).value = prev_wan

    print(f"  [UPD] 商品服务 R{r}: C={old_c}->{cur_wan}, D={old_d}->{fx_wan}, E={old_e}->{prev_wan}  [维修费合计]")
    return 3, 0


def update_maintenance_sheet(wb, data_top_current, data_top_prev, mt_config):
    """更新维修护费 sheet 行4-7的硬编码数值."""
    tl_config = mt_config["top_level"]
    sheet_name = tl_config["sheet"]

    if sheet_name not in [ws.title for ws in wb.worksheets]:
        print(f"[ERROR] Sheet '{sheet_name}' not found")
        return 0, 0

    ws = wb[sheet_name]
    search_col = tl_config["search_col"]
    c_cur = tl_config["data_cols"]["current"] + 1    # C列 (1-indexed)
    c_fx = tl_config["data_cols"]["fengxian"] + 1     # D列
    c_prev = tl_config["data_cols"]["prev"] + 1       # E列

    updated_cells = 0
    skipped = 0

    # Collect all categories including the logistics_entry
    all_categories = list(tl_config["categories"])
    logistics_entry = tl_config.get("logistics_entry")
    if logistics_entry:
        all_categories.append(logistics_entry)

    for cat_def in all_categories:
        label = cat_def["label"]
        expected_row = cat_def.get("expected_row", 0)
        dc = data_top_current.get(label, {})
        dp = data_top_prev.get(label, {})

        cur_total = _round_wan(dc.get("杨浦", 0) + dc.get("奉贤", 0))
        cur_fx = _round_wan(dc.get("奉贤", 0))
        prev_total = _round_wan(dp.get("杨浦", 0) + dp.get("奉贤", 0))

        # Find actual row by label
        actual_row = find_row_by_label(ws, label, search_col, max_row=10)
        if not actual_row:
            actual_row = expected_row
            if actual_row == 0:
                print(f"  [SKIP] 维修护费: '{label}' row not found and no expected_row")
                skipped += 1
                continue

        old_c = ws.cell(actual_row, c_cur).value
        old_d = ws.cell(actual_row, c_fx).value
        old_e = ws.cell(actual_row, c_prev).value

        ws.cell(actual_row, c_cur).value = cur_total
        ws.cell(actual_row, c_fx).value = cur_fx
        ws.cell(actual_row, c_prev).value = prev_total

        print(f"  [UPD] 维修护费 R{actual_row}C: {old_c}->{cur_total} | "
              f"R{actual_row}D: {old_d}->{cur_fx} | R{actual_row}E: {old_e}->{prev_total}  [{label}]")
        updated_cells += 3

    print(f"  [维修护费] Summary: {updated_cells} cells updated, {skipped} rows skipped")
    return updated_cells, skipped


def update_logistics_repair_sheet(wb, data_log_current, data_log_prev, mt_config):
    """更新后勤维修 sheet 左右两个表的硬编码数值."""
    lr_config = mt_config["logistics_repair"]
    sheet_name = lr_config["sheet"]

    if sheet_name not in [ws.title for ws in wb.worksheets]:
        print(f"[ERROR] Sheet '{sheet_name}' not found")
        return 0, 0

    ws = wb[sheet_name]
    updated_cells = 0
    skipped = 0

    # --- Left block ---
    lb = lr_config["left_block"]
    search_col_l = lb["search_col"]
    c_cur_l = lb["data_cols"]["current"] + 1    # D列 (1-indexed)
    c_fx_l = lb["data_cols"]["fengxian"] + 1     # E列
    c_prev_l = lb["data_cols"]["prev"] + 1       # F列

    for sub_cat in lb["sub_categories"]:
        label = sub_cat["label"]
        expected_row = sub_cat.get("expected_row", 0)
        dc = data_log_current.get(label, {})
        dp = data_log_prev.get(label, {})

        cur_total = _round_wan(dc.get("杨浦", 0) + dc.get("奉贤", 0))
        cur_fx = _round_wan(dc.get("奉贤", 0))
        prev_total = _round_wan(dp.get("杨浦", 0) + dp.get("奉贤", 0))

        actual_row = find_row_by_label(ws, label, search_col_l,
                                         max_row=lb["remainder_row"])
        if not actual_row:
            actual_row = expected_row
            if actual_row == 0:
                print(f"  [SKIP] 后勤维修左表: '{label}' row not found")
                skipped += 1
                continue

        old_d = ws.cell(actual_row, c_cur_l).value
        old_e = ws.cell(actual_row, c_fx_l).value
        old_f = ws.cell(actual_row, c_prev_l).value

        ws.cell(actual_row, c_cur_l).value = cur_total
        ws.cell(actual_row, c_fx_l).value = cur_fx
        ws.cell(actual_row, c_prev_l).value = prev_total

        print(f"  [UPD] 后勤维修(左) R{actual_row}D: {old_d}->{cur_total} | "
              f"R{actual_row}E: {old_e}->{cur_fx} | R{actual_row}F: {old_f}->{prev_total}  [{label}]")
        updated_cells += 3

    # --- Right block ---
    rb = lr_config["right_block"]
    search_col_r = rb["search_col"]
    c_cur_r = rb["data_cols"]["current"] + 1    # L列 (1-indexed)
    c_fx_r = rb["data_cols"]["fengxian"] + 1     # M列
    c_prev_r = rb["data_cols"]["prev"] + 1       # N列

    for sub_cat in rb["sub_categories"]:
        sheet_label = sub_cat["label"]            # label used in the sheet
        data_key = sub_cat.get("data_key", sheet_label)  # data key (may differ)
        expected_row = sub_cat.get("expected_row", 0)
        dc = data_log_current.get(data_key, {})
        dp = data_log_prev.get(data_key, {})

        cur_total = _round_wan(dc.get("杨浦", 0) + dc.get("奉贤", 0))
        cur_fx = _round_wan(dc.get("奉贤", 0))
        prev_total = _round_wan(dp.get("杨浦", 0) + dp.get("奉贤", 0))

        actual_row = find_row_by_label(ws, sheet_label, search_col_r,
                                         max_row=rb["remainder_row"])
        if not actual_row:
            actual_row = expected_row
            if actual_row == 0:
                print(f"  [SKIP] 后勤维修右表: '{sheet_label}' row not found")
                skipped += 1
                continue

        old_l = ws.cell(actual_row, c_cur_r).value
        old_m = ws.cell(actual_row, c_fx_r).value
        old_n = ws.cell(actual_row, c_prev_r).value

        ws.cell(actual_row, c_cur_r).value = cur_total
        ws.cell(actual_row, c_fx_r).value = cur_fx
        ws.cell(actual_row, c_prev_r).value = prev_total

        print(f"  [UPD] 后勤维修(右) R{actual_row}L: {old_l}->{cur_total} | "
              f"R{actual_row}M: {old_m}->{cur_fx} | R{actual_row}N: {old_n}->{prev_total}  [{sheet_label}]")
        updated_cells += 3

    print(f"  [后勤维修] Summary: {updated_cells} cells updated, {skipped} rows skipped")
    return updated_cells, skipped


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="新华医院物业管理费及维修费分析（累积月份） — 从累计物业明细更新分析底稿")
    parser.add_argument("current_file", help="本年累计物业及维修明细 Excel 文件（1月至当前月份）")
    parser.add_argument("prev_file", help="上年同期累计物业及维修明细 Excel 文件（1月至当前月份）")
    parser.add_argument("workbook", help="分析底稿 Excel 文件")
    parser.add_argument("--output", "-o", default=None, help="输出路径（默认在底稿名加日期后缀）")
    parser.add_argument("--mapping", "-m", default=None, help="category_mapping.json 路径")
    parser.add_argument("--mode", default="all",
                         choices=["property", "maintenance", "all"],
                         help="分析模式: property(仅物业), maintenance(仅维修), all(全部, 默认)")
    args = parser.parse_args()

    # Validate inputs
    for path, desc in [(args.current_file, "本年明细"),
                        (args.prev_file, "上年明细"),
                        (args.workbook, "分析底稿")]:
        if not os.path.exists(path):
            print(f"[ERROR] {desc} 文件不存在: {path}")
            sys.exit(1)

    # Default mapping path
    if args.mapping is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        args.mapping = os.path.join(script_dir, "..", "references", "category_mapping.json")

    if not os.path.exists(args.mapping):
        print(f"[ERROR] 映射配置文件不存在: {args.mapping}")
        sys.exit(1)

    # Default output
    if args.output is None:
        base, ext = os.path.splitext(args.workbook)
        date_str = datetime.now().strftime("%m%d")
        args.output = f"{base}_updated_{date_str}{ext}"

    # Load mapping config
    with open(args.mapping, 'r', encoding='utf-8') as f:
        config = json.load(f)

    category_map = config["categories"]
    exclude_keywords = config.get("exclude_keywords", [])
    targets = config["targets"]
    mt_config = config.get("maintenance", {})

    print(f"本年明细: {args.current_file}")
    print(f"上年明细: {args.prev_file}")
    print(f"分析底稿: {args.workbook}")
    print(f"输出文件: {args.output}")
    print(f"映射配置: {args.mapping}")
    print(f"分析模式: {args.mode}")

    # -- Load detail files (shared between both analyses) --
    # Load current year
    print(f"\n{'='*50}")
    print("加载本年明细文件...")
    wb_cur, is_xlrd_cur = load_workbook_any(args.current_file)
    ws_cur = get_sheet_by_index(wb_cur, 0, is_xlrd_cur)
    fmt_cur = detect_format(wb_cur, is_xlrd_cur)
    print(f"  检测格式: {'2026新格式' if not fmt_cur.get('has_prefix') else '2025旧格式'}")
    print(f"  文件类型: {'xlrd' if is_xlrd_cur else 'openpyxl'}")

    # Load previous year
    print(f"\n{'='*50}")
    print("加载上年同期明细文件...")
    wb_prev, is_xlrd_prev = load_workbook_any(args.prev_file)
    ws_prev = get_sheet_by_index(wb_prev, 0, is_xlrd_prev)
    fmt_prev = detect_format(wb_prev, is_xlrd_prev)
    print(f"  检测格式: {'2026新格式' if not fmt_prev.get('has_prefix') else '2025旧格式'}")
    print(f"  文件类型: {'xlrd' if is_xlrd_prev else 'openpyxl'}")

    # -- Open analysis workbook once for all updates --
    wb = openpyxl.load_workbook(args.workbook)
    total_updated = 0
    total_skipped = 0

    # --- Property Management Analysis ---
    if args.mode in ("property", "all"):
        print(f"\n{'='*50}")
        print(">>> 物业管理费分析 <<<")

        data_current_pm = load_categories(ws_cur, fmt_cur, category_map,
                                           exclude_keywords, is_xlrd_cur)
        for cat_label in sorted(data_current_pm.keys()):
            campuses = data_current_pm[cat_label]
            details = ", ".join(f"{c}={v:,.0f}元" for c, v in sorted(campuses.items()))
            print(f"  [{cat_label}] {details}")

        data_prev_pm = load_categories(ws_prev, fmt_prev, category_map,
                                        exclude_keywords, is_xlrd_prev)
        for cat_label in sorted(data_prev_pm.keys()):
            campuses = data_prev_pm[cat_label]
            details = ", ".join(f"{c}={v:,.0f}元" for c, v in sorted(campuses.items()))
            print(f"  [{cat_label}] {details}")

        print(f"\n--- 更新物业管理数据 ---")
        upd, sk = update_workbook(wb, data_current_pm, data_prev_pm,
                                   category_map, targets, args.output)
        total_updated += upd
        total_skipped += sk

    # --- Maintenance Fee Analysis ---
    if args.mode in ("maintenance", "all"):
        print(f"\n{'='*50}")
        print(">>> 维修费分析 <<<")

        if not mt_config:
            print("[ERROR] maintenance 配置缺失，请在 category_mapping.json 中添加 maintenance 段")
        else:
            data_current_mt = load_maintenance_data(ws_cur, fmt_cur, is_xlrd_cur)
            data_prev_mt = load_maintenance_data(ws_prev, fmt_prev, is_xlrd_prev)

            # Print top-level summary
            print("\n  一级分类（维修护费）:")
            for cat_label in sorted(data_current_mt["top_level"].keys()):
                campuses = data_current_mt["top_level"][cat_label]
                details = ", ".join(f"{c}={v:,.0f}元" for c, v in sorted(campuses.items()))
                total = sum(campuses.values())
                print(f"  [{cat_label}] {details} | 合计={total:,.0f}元 = {_round_wan(total)}万元")

            print("\n  二级分类（后勤维修）:")
            for cat_label in sorted(data_current_mt["logistics_repair"].keys()):
                campuses = data_current_mt["logistics_repair"][cat_label]
                details = ", ".join(f"{c}={v:,.0f}元" for c, v in sorted(campuses.items()))
                total = sum(campuses.values())
                print(f"  [{cat_label}] {details} | 合计={total:,.0f}元 = {_round_wan(total)}万元")

            # Cross-validation
            logistics_sum = sum(
                sum(d.values()) for d in data_current_mt["logistics_repair"].values()
            )
            logistics_top = sum(
                data_current_mt["top_level"].get("后勤维修", {}).values()
            )
            print(f"\n  [验证] 后勤维修子项之和={logistics_sum:,.0f}元 vs 一级后勤维修={logistics_top:,.0f}元")

            # Update sheets
            print(f"\n--- 更新维修费数据 ---")
            upd1, sk1 = update_service_row_for_maintenance(
                wb, data_current_mt["top_level"], data_prev_mt["top_level"], mt_config)
            upd2, sk2 = update_maintenance_sheet(
                wb, data_current_mt["top_level"], data_prev_mt["top_level"], mt_config)
            upd3, sk3 = update_logistics_repair_sheet(
                wb, data_current_mt["logistics_repair"], data_prev_mt["logistics_repair"], mt_config)
            total_updated += upd1 + upd2 + upd3
            total_skipped += sk1 + sk2 + sk3

    # -- Close detail workbooks --
    if is_xlrd_cur:
        wb_cur.release_resources()
    else:
        wb_cur.close()
    if is_xlrd_prev:
        wb_prev.release_resources()
    else:
        wb_prev.close()

    # -- Save analysis workbook --
    wb.save(args.output)
    wb.close()

    print(f"\n{'='*50}")
    print(f"All done! Total: {total_updated} cells updated, {total_skipped} rows skipped")
    print(f"Output: {args.output}")


if __name__ == "__main__":
    main()
