#!/usr/bin/env python3
"""
新华医院成本分析模块 PPT 数据更新工具。

用法:
    python cost_analysis.py <excel_path> <pptx_path> [--output OUTPUT]
    python cost_analysis.py 4月分析底稿.xlsx 新华医院2026年4月运营分析0509.pptx

从 Excel 附表和数据表中提取关键成本/收入数据，除以 10000 后更新到 PPT 成本分析模块。
"""

import json
import os
import sys
import re
import argparse
from datetime import datetime
from collections import defaultdict

import openpyxl
from pptx import Presentation


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def detect_data_month(excel_path):
    """Extract month number from filename like '4月分析底稿.xlsx' → 4."""
    basename = os.path.basename(excel_path)
    m = re.search(r'(\d{1,2})\s*月', basename)
    if m:
        return int(m.group(1))
    return 4  # default


# ---------------------------------------------------------------------------
# ExcelDataLoader
# ---------------------------------------------------------------------------

class ExcelDataLoader:
    """加载 Excel 所有 sheet，建立 {source_alias: {normalized_label: {col_idx: value}}} 索引."""

    def __init__(self, excel_path, sources_config):
        self.excel_path = excel_path
        self.sources = sources_config
        self.cache = {}       # source_alias -> {label -> {col: value}}
        self.needs_wan = {}   # source_alias -> bool
        self.label_col = {}   # source_alias -> int
        self._load_all()

    @staticmethod
    def _norm(s):
        """标准化文本用于匹配."""
        if s is None:
            return ""
        s = str(s).strip()
        s = re.sub(r'[\s　\n\r\t]+', '', s)
        return s

    def _match_sheet(self, wb, target_name):
        """通过精确匹配 → 标准化匹配 → 包含匹配找 sheet."""
        # Exact match
        for ws in wb.worksheets:
            if ws.title == target_name:
                return ws
        # Normalized match
        tn = self._norm(target_name)
        for ws in wb.worksheets:
            if self._norm(ws.title) == tn:
                return ws
        # Substring either way
        for ws in wb.worksheets:
            wsn = self._norm(ws.title)
            if tn and (tn in wsn or wsn in tn):
                return ws
        return None

    def _load_all(self):
        wb = openpyxl.load_workbook(self.excel_path, data_only=True)
        available = [ws.title for ws in wb.worksheets]
        print(f"  Excel sheets ({len(available)}): {available}")

        for alias, cfg in self.sources.items():
            sheet_name = cfg["sheet"]
            needs = cfg.get("needs_wan", False)
            self.needs_wan[alias] = needs
            label_col = cfg.get("label_col", 0)
            self.label_col[alias] = label_col

            ws = self._match_sheet(wb, sheet_name)
            if ws is None:
                print(f"  [WARN] Sheet not found: '{sheet_name}'")
                self.cache[alias] = {}
                continue

            sheet_data = {}
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                label = self._norm(row[label_col].value)
                if label:
                    row_dict = {}
                    for c, cell in enumerate(row):
                        row_dict[c] = cell.value
                    sheet_data[label] = row_dict
            self.cache[alias] = sheet_data
            print(f"  [LOAD] {alias} ← {ws.title} ({len(sheet_data)} rows)")
        wb.close()
        self._repair_cache()

    def _fuzzy_get(self, sheet_data, target_label):
        """模糊匹配：exact > contains > startswith."""
        if target_label in sheet_data:
            return sheet_data[target_label]
        for label in sheet_data:
            if target_label in label:
                return sheet_data[label]
        for label in sheet_data:
            if label in target_label:
                return sheet_data[label]
        return None

    def _has_data(self, alias):
        """Check if a cached source has any non-None numeric values."""
        if alias not in self.cache:
            return False
        for label, row in self.cache[alias].items():
            for k, v in row.items():
                if isinstance(v, (int, float)) and v is not None:
                    return True
        return False

    def _get_from_cache(self, alias, label, col):
        """Get a raw value from another cache entry (bypasses wan conversion)."""
        if alias not in self.cache:
            return None
        row = self._fuzzy_get(self.cache[alias], self._norm(label))
        if row is None:
            return None
        actual_col = col + self.label_col.get(alias, 0)
        val = row.get(actual_col)
        if not isinstance(val, (int, float)):
            return None
        if self.needs_wan.get(alias, False):
            val = val / 10000.0
        return val

    def _get_from_cache_no_wan(self, alias, label, col):
        """Get raw value without wan conversion."""
        if alias not in self.cache:
            return None
        row = self._fuzzy_get(self.cache[alias], self._norm(label))
        if row is None:
            return None
        actual_col = col + self.label_col.get(alias, 0)
        val = row.get(actual_col)
        if not isinstance(val, (int, float)):
            return None
        return val

    def _repair_cache(self):
        """Repair formula-sheet caches whose values are empty due to unrecalculated formulas."""
        self._repair_medical_cost()
        self._repair_personnel()
        self._repair_energy()
        self._repair_supplementary3()         # must run before _repair_services
        self._repair_supplementary3_material()
        self._repair_services()
        self._repair_surplus()
        self._repair_assay_income()

    def _is_row_empty(self, cache, label, cols_to_check):
        """Check if specific columns in a cached row are all None."""
        row = self._fuzzy_get(cache, label)
        if row is None:
            return True
        for c in cols_to_check:
            if isinstance(row.get(c), (int, float)):
                return False
        return True

    def _repair_medical_cost(self):
        """Repair 医疗成本 sheet from 附表2 sources."""
        cache = self.cache.get("医疗成本", {})
        if not cache:
            return
        # Check if key data rows are empty
        if not self._is_row_empty(cache, "人员经费", [2, 4, 6]):
            return
        print("  [REPAIR] 医疗成本: populating from 附表2 sources...")
        cache = self.cache.get("医疗成本", {})
        if not cache:
            return

        row_map = {
            "人员经费": "1.人员经费",
            "药品费": "3.药品费",
            "卫生材料费": "2.卫生材料费",
            "医疗风险基金": "6.提取医疗风险基金",
            "商品和服务费": "7.其他费用",
            "医疗成本合计": "一、医疗成本",
        }

        for label, xlbl in row_map.items():
            row = self._fuzzy_get(cache, label)
            if row is None:
                continue
            # Col 2: 当月医院, Col 4: 奉贤, Col 6: 同期
            cur = self._get_from_cache("当月附表2-成本", xlbl, 1)
            fx = self._get_from_cache("当月奉贤附表2-成本", xlbl, 1)
            prev = self._get_from_cache("上年同期附表2-成本", xlbl, 1)
            if cur is not None:
                row[2] = cur
            if fx is not None:
                row[4] = fx
            if prev is not None:
                row[6] = prev
            # Col 3/5/7: 占比%
            if cur is not None:
                total_cur = self._get_from_cache("当月附表2-成本", "一、医疗成本", 1)
                if total_cur and total_cur != 0:
                    row[3] = cur / total_cur
            if fx is not None:
                total_fx = self._get_from_cache("当月奉贤附表2-成本", "一、医疗成本", 1)
                if total_fx and total_fx != 0:
                    row[5] = fx / total_fx
            if prev is not None:
                total_prev = self._get_from_cache("上年同期附表2-成本", "一、医疗成本", 1)
                if total_prev and total_prev != 0:
                    row[7] = prev / total_prev
            # Col 8: 增减额 = cur - prev
            if cur is not None and prev is not None:
                row[8] = cur - prev
            # Col 9: 增减率 = (cur - prev) / |prev|
            if cur is not None and prev is not None and prev != 0:
                row[9] = (cur - prev) / abs(prev)

        # Special: 折旧/摊销 = 固定资产折旧费 + 无形资产摊销费
        row = self._fuzzy_get(cache, "折旧/摊销")
        if row is not None:
            for src_name, col_idx in [("当月附表2-成本", 2), ("当月奉贤附表2-成本", 4), ("上年同期附表2-成本", 6)]:
                d1 = self._get_from_cache(src_name, "4.固定资产折旧费", 1)
                d2 = self._get_from_cache(src_name, "5.无形资产摊销费", 1)
                if d1 is not None and d2 is not None:
                    row[col_idx] = d1 + d2
            # Compute derived columns for 折旧/摊销
            cur = row.get(2)
            prev = row.get(6)
            if cur is not None:
                total_cur = self._get_from_cache("当月附表2-成本", "一、医疗成本", 1)
                if total_cur and total_cur != 0:
                    row[3] = cur / total_cur
            if row.get(4) is not None:
                total_fx = self._get_from_cache("当月奉贤附表2-成本", "一、医疗成本", 1)
                if total_fx and total_fx != 0:
                    row[5] = row[4] / total_fx
            if prev is not None:
                total_prev = self._get_from_cache("上年同期附表2-成本", "一、医疗成本", 1)
                if total_prev and total_prev != 0:
                    row[7] = prev / total_prev
            if cur is not None and prev is not None:
                row[8] = cur - prev
                if prev != 0:
                    row[9] = (cur - prev) / abs(prev)

        print(f"  [REPAIR] 医疗成本: done")

    def _repair_personnel(self):
        """Repair 人员经费 sheet from 附表3 sources."""
        cache = self.cache.get("人员经费", {})
        if not cache:
            return
        if not self._is_row_empty(cache, "人员经费", [2, 3, 4]):
            return
        print("  [REPAIR] 人员经费: populating from 附表3 sources...")
        cache = self.cache.get("人员经费", {})
        if not cache:
            return

        # 医疗服务收入 → 一、医疗收入 from 附表2 (income side)
        row = self._fuzzy_get(cache, "医疗服务收入")
        if row is not None:
            for src, lbl, col_idx in [("当月附表2", "一、医疗收入", 2), ("当月奉贤附表2", "一、医疗收入", 3),
                                       ("上年同期附表2", "一、医疗收入", 4)]:
                v = self._get_from_cache(src, lbl, 1)
                if v is not None:
                    row[col_idx] = v

        # 人员经费
        row = self._fuzzy_get(cache, "人员经费")
        if row is not None:
            cur = self._get_from_cache("当月附表3", "（一）人员经费", 1)
            fx = self._get_from_cache("当月奉贤附表3", "（一）人员经费", 1)
            prev = self._get_from_cache("上年同期附表2-成本", "1.人员经费", 1)
            if cur is not None:
                row[2] = cur
            if fx is not None:
                row[3] = fx
            if prev is not None:
                row[4] = prev

        # 工资总额 = sum of sub-components
        wage_parts = ["基本工资", "津贴补贴", "伙食补助费", "绩效工资", "其他工资福利支出"]
        row = self._fuzzy_get(cache, "其中：工资总额")
        if row is not None:
            cur_total = sum(
                (self._get_from_cache_no_wan("当月附表3", p, 1) or 0) for p in wage_parts) / 10000.0
            fx_total = sum(
                (self._get_from_cache_no_wan("当月奉贤附表3", p, 1) or 0) for p in wage_parts) / 10000.0
            prev_total = sum(
                (self._get_from_cache_no_wan("上年同期附表3", p, 3) or 0) +
                (self._get_from_cache_no_wan("上年同期附表3", p, 5) or 0)
                for p in wage_parts) / 10000.0
            row[2] = cur_total
            row[3] = fx_total
            row[4] = prev_total

        # 基本工资等 = sum of basic+allowance+food+other (no 绩效工资)
        base_parts = ["基本工资", "津贴补贴", "伙食补助费", "其他工资福利支出"]
        row = self._fuzzy_get(cache, "其中：基本工资等")
        if row is not None:
            cur_total = sum(
                (self._get_from_cache_no_wan("当月附表3", p, 1) or 0) for p in base_parts) / 10000.0
            fx_total = sum(
                (self._get_from_cache_no_wan("当月奉贤附表3", p, 1) or 0) for p in base_parts) / 10000.0
            prev_total = sum(
                (self._get_from_cache_no_wan("上年同期附表3", p, 3) or 0) +
                (self._get_from_cache_no_wan("上年同期附表3", p, 5) or 0)
                for p in base_parts) / 10000.0
            row[2] = cur_total
            row[3] = fx_total
            row[4] = prev_total

        # 绩效工资
        row = self._fuzzy_get(cache, "绩效工资")
        if row is not None:
            cur = self._get_from_cache("当月附表3", "绩效工资", 1)
            fx = self._get_from_cache("当月奉贤附表3", "绩效工资", 1)
            prev = (self._get_from_cache_no_wan("上年同期附表3", "绩效工资", 3) or 0) + \
                   (self._get_from_cache_no_wan("上年同期附表3", "绩效工资", 5) or 0)
            prev = prev / 10000.0
            if cur is not None:
                row[2] = cur
            if fx is not None:
                row[3] = fx
            if prev is not None:
                row[4] = prev

        # 社保公积金 = sum of insurance + housing fund
        insurance_parts = ["基本养老保险缴费", "职业年金缴费", "基本医疗保险缴费",
                          "其他社会保障缴费", "住房公积金"]
        row = self._fuzzy_get(cache, "社保公积金")
        if row is not None:
            cur_total = sum(
                (self._get_from_cache_no_wan("当月附表3", p, 1) or 0) for p in insurance_parts) / 10000.0
            fx_total = sum(
                (self._get_from_cache_no_wan("当月奉贤附表3", p, 1) or 0) for p in insurance_parts) / 10000.0
            prev_total = sum(
                (self._get_from_cache_no_wan("上年同期附表3", p, 3) or 0) +
                (self._get_from_cache_no_wan("上年同期附表3", p, 5) or 0)
                for p in insurance_parts) / 10000.0
            row[2] = cur_total
            row[3] = fx_total
            row[4] = prev_total

        # Compute 增减额 and 增减率 for all rows
        for label in ["医疗服务收入", "人员经费", "其中：工资总额", "其中：基本工资等", "绩效工资", "社保公积金"]:
            row = self._fuzzy_get(cache, label)
            if row is None:
                continue
            cur = row.get(2)
            prev = row.get(4)
            if cur is not None and prev is not None:
                row[5] = cur - prev
                if prev != 0:
                    row[6] = (cur - prev) / abs(prev)

        print(f"  [REPAIR] 人员经费: done")

    def _repair_energy(self):
        """Repair 水电煤 sheet from 附表3 sources."""
        cache = self.cache.get("水电煤", {})
        if not cache:
            return
        # Run if any key data column is empty (check individually, not all-or-nothing)
        need_cur = self._is_row_empty(cache, "电费", [2])
        need_fx = self._is_row_empty(cache, "电费", [3])
        need_prev = self._is_row_empty(cache, "电费", [4])
        if not need_cur and not need_fx and not need_prev:
            return
        print("  [REPAIR] 水电煤: populating from 附表3 sources...")
        cache = self.cache.get("水电煤", {})
        if not cache:
            return

        energy_items = {
            "其中：水费": "水费",
            "电费": "电费",
            "燃气费": "燃气费",
        }

        # Populate individual energy items
        for label, xlbl in energy_items.items():
            row = self._fuzzy_get(cache, label)
            if row is None:
                continue
            # Col 2: 当月累计医院, Col 3: 奉贤累计, Col 4: 同期累计
            for src, src_label, col_idx in [("当月附表3", xlbl, 2), ("当月奉贤附表3", xlbl, 3),
                                              ("上年同期附表3", xlbl, 4)]:
                if col_idx == 4:
                    # 上年同期附表3: sum of 业务活动费用(col 4) + 单位管理费用(col 6)
                    v1 = self._get_from_cache_no_wan(src, src_label, 4)
                    v2 = self._get_from_cache_no_wan(src, src_label, 6)
                    v = ((v1 or 0) + (v2 or 0)) / 10000.0
                else:
                    # Use 累计数 (col 2 for total, or col 4 for 业务活动+col 6 for 单位管理)
                    v = self._get_from_cache(src, src_label, 2)
                    if v is None and col_idx == 2:
                        v1 = self._get_from_cache_no_wan(src, src_label, 4)
                        v2 = self._get_from_cache_no_wan(src, src_label, 6)
                        v = ((v1 or 0) + (v2 or 0)) / 10000.0
                if v is not None:
                    row[col_idx] = v

        # 能耗成本 = sum of sub-items
        row = self._fuzzy_get(cache, "能耗成本")
        if row is not None:
            sub_labels = ["其中：水费", "电费", "燃气费"]
            for col_idx in [2, 3, 4]:
                total = 0
                for sl in sub_labels:
                    sr = self._fuzzy_get(cache, sl)
                    if sr and isinstance(sr.get(col_idx), (int, float)):
                        total += sr[col_idx]
                if total > 0:
                    row[col_idx] = total

        # Compute 增减额 and 增减率
        for label in ["能耗成本", "其中：水费", "电费", "燃气费"]:
            row = self._fuzzy_get(cache, label)
            if row is None:
                continue
            cur = row.get(2)
            prev = row.get(4)
            if cur is not None and prev is not None:
                row[5] = cur - prev  # 增减额
                if prev != 0:
                    row[6] = (cur - prev) / abs(prev)  # 增减率

        print(f"  [REPAIR] 水电煤: done")

    def _repair_services(self):
        """Repair 商品服务, 维修护费 sheets whose summary rows are empty."""
        self._repair_goods_services()
        self._repair_maintenance()
        self._repair_derived_columns()

    def _repair_goods_services(self):
        """Repair 商品服务 sheet from 附表3 data.

        The 商品服务 sheet has several rows with empty data (能源成本, 委托业务费,
        劳务费, 租赁费, 其他). These must be populated from 附表3 for the corresponding
        time periods (当月→当月附表3, 奉贤→当月奉贤附表3, 上年同期→上年同期附表3).

        商品服务 sheet layout (label_col=1): col2=当月, col3=奉贤, col4=上年同期,
                                              col5=增减额, col6=增减率.
        附表3 col layout: col1=合计本月数 (label_col=0, needs_wan).
        """
        cache = self.cache.get("商品服务", {})
        if not cache:
            return

        # Check which rows need repair
        need_total = self._is_row_empty(cache, "商品服务", [2, 3, 4])
        need_energy = self._is_row_empty(cache, "其中：能源成本", [2, 3, 4])
        need_commission = self._is_row_empty(cache, "委托业务费", [2, 3, 4])
        need_labor = self._is_row_empty(cache, "劳务费", [2, 3, 4])
        need_rental = self._is_row_empty(cache, "租赁费", [2, 3, 4])
        need_other = self._is_row_empty(cache, "其他", [2, 3, 4])

        if not any([need_total, need_energy, need_commission, need_labor,
                    need_rental, need_other]):
            return

        print("  [REPAIR] 商品服务: populating from 附表3...")

        # Helper: read value from 附表3 and convert to 万元
        def _read_s3(s3_alias, label, col=1):
            """Read a value from a 附表3 cache, already in 万元 (needs_wan=true source)."""
            val = self._get_from_cache_no_wan(s3_alias, label, col)
            if isinstance(val, (int, float)):
                return val / 10000.0
            return None

        # Period → (附表3 alias, 商品服务 col)
        periods = [
            ("当月附表3", 2),       # 当月 → col 2
            ("当月奉贤附表3", 3),   # 奉贤 → col 3
            ("上年同期附表3", 4),   # 上年同期 → col 4
        ]

        # Energy cost items (water + electricity + gas from 附表3)
        energy_items = ["水费", "电费", "燃气费"]

        # --- Populate 商品服务 total (row 54: （七）商品和服务费用) ---
        # Label uses fullwidth parens （）in Excel, so search by core text
        if need_total:
            row = self._fuzzy_get(cache, "商品服务")
            if row is not None:
                for s3_alias, goods_col in periods:
                    v = _read_s3(s3_alias, "商品和服务费用", 1)
                    if v is not None:
                        row[goods_col] = v
                self._compute_derived(row, 2, 4, 5, 6)
                print(f"  [REPAIR] 商品服务 total: cur={row.get(2)}, fx={row.get(3)}, "
                      f"prev={row.get(4)}")

        # --- Populate 能源成本 (sum of utility items from 附表3) ---
        if need_energy:
            row = self._fuzzy_get(cache, "其中：能源成本")
            if row is not None:
                for s3_alias, goods_col in periods:
                    total = 0
                    for item in energy_items:
                        v = _read_s3(s3_alias, item, 1)
                        if v is not None:
                            total += v
                    row[goods_col] = total
                self._compute_derived(row, 2, 4, 5, 6)
                print(f"  [REPAIR] 能源成本: cur={row.get(2):.2f}, fx={row.get(3):.2f}, "
                      f"prev={row.get(4):.2f}")

        # --- Populate 委托业务费 (附表3 Row 74) ---
        if need_commission:
            row = self._fuzzy_get(cache, "委托业务费")
            if row is not None:
                for s3_alias, goods_col in periods:
                    v = _read_s3(s3_alias, "委托业务费", 1)
                    if v is not None:
                        row[goods_col] = v
                self._compute_derived(row, 2, 4, 5, 6)
                print(f"  [REPAIR] 委托业务费: cur={row.get(2):.2f}, fx={row.get(3):.2f}, "
                      f"prev={row.get(4):.2f}")

        # --- Populate 劳务费 (附表3 Row 73) ---
        if need_labor:
            row = self._fuzzy_get(cache, "劳务费")
            if row is not None:
                for s3_alias, goods_col in periods:
                    v = _read_s3(s3_alias, "劳务费", 1)
                    if v is not None:
                        row[goods_col] = v
                self._compute_derived(row, 2, 4, 5, 6)
                print(f"  [REPAIR] 劳务费: cur={row.get(2):.2f}, fx={row.get(3):.2f}, "
                      f"prev={row.get(4):.2f}")

        # --- Populate 租赁费 (附表3 Row 67) ---
        if need_rental:
            row = self._fuzzy_get(cache, "租赁费")
            if row is not None:
                for s3_alias, goods_col in periods:
                    v = _read_s3(s3_alias, "租赁费", 1)
                    if v is not None:
                        row[goods_col] = v
                self._compute_derived(row, 2, 4, 5, 6)
                print(f"  [REPAIR] 租赁费: cur={row.get(2):.2f}, fx={row.get(3):.2f}, "
                      f"prev={row.get(4):.2f}")

        # --- Populate 其他 = 商品服务 - Σ(explicit items) ---
        # Explicit items on Slide 54: 能源成本, 维修护费, 物业管理费, 委托业务费,
        # 劳务费, 租赁费.  Also subtract 专用材料费 and 供应商履约成本 which are
        # in the 商品服务 sheet but not shown on Slide 54 — they roll into 其他.
        # Always recompute to ensure consistency with corrected 能源成本.
        if True:  # was: if need_other — always recompute for consistency
            row = self._fuzzy_get(cache, "其他")
            total_row = self._fuzzy_get(cache, "商品服务")
            if row is not None and total_row is not None:
                sub_labels = [
                    "其中：能源成本", "维修护费", "物业管理费",
                    "委托业务费", "劳务费", "专用材料费",
                    "供应商履约成本", "租赁费",
                ]
                for goods_col in [2, 3, 4]:
                    total_val = total_row.get(goods_col)
                    if isinstance(total_val, (int, float)):
                        subtotal = 0
                        for sl in sub_labels:
                            sr = self._fuzzy_get(cache, sl)
                            if sr is not None:
                                sv = sr.get(goods_col)
                                if isinstance(sv, (int, float)):
                                    subtotal += sv
                        row[goods_col] = total_val - subtotal
                    else:
                        row[goods_col] = 0
                self._compute_derived(row, 2, 4, 5, 6)
                print(f"  [REPAIR] 其他: cur={row.get(2):.2f}, fx={row.get(3):.2f}, "
                      f"prev={row.get(4):.2f}")

        print(f"  [REPAIR] 商品服务: done")

    @staticmethod
    def _compute_derived(row, cur_col, prev_col, diff_col, rate_col):
        """Compute 增减额 and 增减率 for a row."""
        c = row.get(cur_col)
        p = row.get(prev_col)
        if isinstance(c, (int, float)) and isinstance(p, (int, float)):
            row[diff_col] = c - p
            if p != 0:
                row[rate_col] = (c - p) / abs(p)

    def _repair_maintenance(self):
        """Repair 维修护费 sheet.

        '维修费' total row is empty → sum of indented sub-items.
        Layout (label_col=1): idx 1=label, 2=当月, 3=奉贤, 4=同期, 5=增减额, 6=增减率.
        """
        cache = self.cache.get("维修护费", {})
        if not cache:
            return

        if not self._is_row_empty(cache, "维修费", [2, 3, 4]):
            return

        print("  [REPAIR] 维修护费: populating from sub-items...")

        sub_labels = [
            "网络维修",
            "设备维修",
            "后勤维修",
            "房屋维修",
        ]

        row = self._fuzzy_get(cache, "维修费")
        if row is not None:
            for col_idx in [2, 3, 4]:  # 当月, 奉贤, 同期
                total = 0
                has_data = False
                for sl in sub_labels:
                    sr = self._fuzzy_get(cache, sl)
                    if sr and isinstance(sr.get(col_idx), (int, float)):
                        total += sr[col_idx]
                        has_data = True
                if has_data:
                    row[col_idx] = total

            # Derived columns
            cur = row.get(2)
            prev = row.get(4)
            if isinstance(cur, (int, float)) and isinstance(prev, (int, float)):
                row[5] = cur - prev  # 增减额
                if prev != 0:
                    row[6] = (cur - prev) / abs(prev)  # 增减率

        print(f"  [REPAIR] 维修护费: done")

    def _repair_derived_columns(self):
        """Compute missing 增减额 and 增减率 for rows that have 当月/同期 data.

        Many front-end Excel sheets have raw data in the 当月/奉贤/同期 columns
        but leave the 增减额/增减率 formula columns uncalculated. This repair
        fills in those derived columns for all affected sheets.
        """
        # (cache_key, cur_col, fx_col, prev_col, diff_col, rate_col)
        sheet_layouts = [
            ("商品服务",  2, 3, 4, 5, 6),
            ("物业管理",  9, 10, 11, 12, 13),
            ("维修护费",  2, 3, 4, 5, 6),
        ]

        for alias, cur_c, fx_c, prev_c, diff_c, rate_c in sheet_layouts:
            cache = self.cache.get(alias, {})
            if not cache:
                continue
            repaired = 0
            for label, row in cache.items():
                cur = row.get(cur_c)
                prev = row.get(prev_c)
                diff = row.get(diff_c)
                rate = row.get(rate_c)

                if not isinstance(cur, (int, float)) or not isinstance(prev, (int, float)):
                    continue

                changed = False
                if not isinstance(diff, (int, float)):
                    row[diff_c] = cur - prev
                    changed = True
                if not isinstance(rate, (int, float)) and prev != 0:
                    row[rate_c] = (cur - prev) / abs(prev)
                    changed = True
                if changed:
                    repaired += 1

            if repaired > 0:
                print(f"  [REPAIR] {alias}: computed derived columns for {repaired} rows")

    def _repair_surplus(self):
        """Repair 上年同期附表1 - compute 医疗业务盈余/医疗服务盈余/其他盈余
        from sub-items when the formula cells are empty (uncalculated)."""
        cache = self.cache.get("上年同期附表1", {})
        if not cache:
            return
        # Only repair if 医疗业务盈余 cells are empty
        if not self._is_row_empty(cache, "一、医疗业务盈余：", [1, 2]):
            return
        print("  [REPAIR] 上年同期附表1: computing surplus from sub-items...")

        # --- 当月值 (col 1) ---
        med_income_cur = self._get_from_cache("上年同期附表1", "医疗收入", 1)
        fiscal_cur = self._get_from_cache("上年同期附表1", "加：财政基本拨款收入", 1)
        med_cost_cur = self._get_from_cache("上年同期附表1", "减：医疗成本", 1)
        other_income_cur = self._get_from_cache("上年同期附表1", "其他医疗活动收入", 1)
        other_expense_cur = self._get_from_cache("上年同期附表1", "减：其他医疗活动费用", 1)

        # --- 累计值 (col 2) ---
        med_income_acc = self._get_from_cache("上年同期附表1", "医疗收入", 2)
        fiscal_acc = self._get_from_cache("上年同期附表1", "加：财政基本拨款收入", 2)
        med_cost_acc = self._get_from_cache("上年同期附表1", "减：医疗成本", 2)
        other_income_acc = self._get_from_cache("上年同期附表1", "其他医疗活动收入", 2)
        other_expense_acc = self._get_from_cache("上年同期附表1", "减：其他医疗活动费用", 2)

        # 医疗服务盈余 = 医疗收入 + 财政拨款 - 医疗成本
        svc_cur = None
        svc_acc = None
        if all(v is not None for v in [med_income_cur, fiscal_cur, med_cost_cur]):
            svc_cur = med_income_cur + fiscal_cur - med_cost_cur
        if all(v is not None for v in [med_income_acc, fiscal_acc, med_cost_acc]):
            svc_acc = med_income_acc + fiscal_acc - med_cost_acc

        # 其他盈余 = 其他收入 - 其他费用
        other_cur = None
        other_acc = None
        if all(v is not None for v in [other_income_cur, other_expense_cur]):
            other_cur = other_income_cur - other_expense_cur
        if all(v is not None for v in [other_income_acc, other_expense_acc]):
            other_acc = other_income_acc - other_expense_acc

        # 医疗业务盈余 = 医疗服务盈余 + 其他盈余
        biz_cur = None
        biz_acc = None
        if svc_cur is not None and other_cur is not None:
            biz_cur = svc_cur + other_cur
        if svc_acc is not None and other_acc is not None:
            biz_acc = svc_acc + other_acc

        # Write to cache: row 5 (一、医疗业务盈余), row 6 (医疗服务盈余), row 11 (其他盈余)
        # Note: _get_from_cache divides by 10000 for needs_wan sheets, so computed
        # values are in 万元. Multiply back to 元 for cache storage so that
        # get_value (which also divides by 10000) returns correct 万元 values.
        for label, cur_val, acc_val in [
            ("一、医疗业务盈余：", biz_cur, biz_acc),
            ("（一）医疗服务盈余：", svc_cur, svc_acc),
            ("（二）其他盈余：", other_cur, other_acc),
        ]:
            row = self._fuzzy_get(cache, label)
            if row is None:
                continue
            if cur_val is not None:
                row[1] = cur_val * 10000
            if acc_val is not None:
                row[2] = acc_val * 10000

        print(f"  [REPAIR] 上年同期附表1: done")

    def _repair_supplementary3(self):
        """Repair 上年同期附表3 - compute 合计 columns from 业务活动费用 + 单位管理费用
        when the formula cells are empty (uncalculated).

        附表3 layout: col1=合计本月数, col2=合计累计数,
                       col3=业务活动费本月数, col4=业务活动费累计数,
                       col5=单位管理费本月数, col6=单位管理费累计数
        """
        cache = self.cache.get("上年同期附表3", {})
        if not cache:
            return
        repaired = 0
        for label, row in cache.items():
            if not isinstance(row, dict):
                continue
            col1 = row.get(1)  # 合计本月数
            col2 = row.get(2)  # 合计累计数
            col3 = row.get(3)  # 业务活动费本月数
            col4 = row.get(4)  # 业务活动费累计数
            col5 = row.get(5)  # 单位管理费本月数
            col6 = row.get(6)  # 单位管理费累计数

            changed = False
            if not isinstance(col1, (int, float)):
                v3 = col3 if isinstance(col3, (int, float)) else 0
                v5 = col5 if isinstance(col5, (int, float)) else 0
                if v3 != 0 or v5 != 0:
                    row[1] = v3 + v5
                    changed = True
            if not isinstance(col2, (int, float)):
                v4 = col4 if isinstance(col4, (int, float)) else 0
                v6 = col6 if isinstance(col6, (int, float)) else 0
                if v4 != 0 or v6 != 0:
                    row[2] = v4 + v6
                    changed = True
            if changed:
                repaired += 1

        if repaired > 0:
            print(f"  [REPAIR] 上年同期附表3: computed 合计 for {repaired} rows")

        # --- Repair 七、商品和服务费用 total row (fully empty → sum sub-items) ---
        # Label uses fullwidth parens in Excel: （七）商品和服务费用
        goods_row = self._fuzzy_get(cache, "商品和服务费用")
        if goods_row is not None:
            col1 = goods_row.get(1)
            col2 = goods_row.get(2)
            if not isinstance(col1, (int, float)) or not isinstance(col2, (int, float)):
                total_cur = 0
                total_acc = 0
                goods_section = False
                # Iterate in insertion order (= Excel row order), NOT sorted.
                # （七）starts with fullwidth paren U+FF08 which sorts after CJK chars.
                for label, row in cache.items():
                    if not isinstance(row, dict):
                        continue
                    if '商品和服务费用' in label:
                        goods_section = True
                        continue
                    if goods_section:
                        if re.match(r'^[八九十]、', label):
                            break
                        c1 = row.get(1)
                        c2 = row.get(2)
                        if isinstance(c1, (int, float)):
                            total_cur += c1
                        if isinstance(c2, (int, float)):
                            total_acc += c2
                if total_cur != 0 or total_acc != 0:
                    goods_row[1] = total_cur
                    goods_row[2] = total_acc
                    print(f"  [REPAIR] 上年同期附表3 七、商品和服务费用: summed from sub-items "
                          f"(cur={total_cur:.2f}, acc={total_acc:.2f})")

    def _repair_supplementary3_material(self):
        """Repair 上年同期附表3 卫生材料费 from 上年同期附表2-成本.

        The 附表3 repair may under-count 卫生材料费 when the 业务活动费 columns
        are empty (only 单位管理费 has data). Cross-reference from 附表2-成本
        which has the correct full value for this specific row.
        """
        cache = self.cache.get("上年同期附表3", {})
        if not cache:
            return
        row = self._fuzzy_get(cache, "卫生材料费")
        if row is None:
            return
        # If 合计本月数 already has a reasonable value, skip
        col1 = row.get(1)  # 合计本月数
        col2 = row.get(2)  # 合计累计数
        # Get correct values from 附表2-成本
        val_cur = self._get_from_cache_no_wan("上年同期附表2-成本", "2.卫生材料费", 1)
        val_acc = self._get_from_cache_no_wan("上年同期附表2-成本", "2.卫生材料费", 2)
        changed = False
        if val_cur is not None and val_cur > 0:
            current_cur = col1 if isinstance(col1, (int, float)) else 0
            # Only repair if the current value is clearly too small (< 20% of expected)
            if current_cur < val_cur * 0.2:
                row[1] = val_cur
                changed = True
        if val_acc is not None and val_acc > 0:
            current_acc = col2 if isinstance(col2, (int, float)) else 0
            if current_acc < val_acc * 0.2:
                row[2] = val_acc
                changed = True
        if changed:
            print(f"  [REPAIR] 上年同期附表3 卫生材料费: populated from 附表2-成本 "
                  f"(cur={val_cur}, acc={val_acc})")

    def _repair_assay_income(self):
        """Repair 化验收入 in 附表2 - sum 门急诊 + 住院 occurrences.

        附表2 收入侧有两个"化验收入"行：门急诊收入下的化验收入和住院收入下的化验收入。
        加载时后者覆盖前者，导致数据不完整。此修复重新读取原始 Excel，
        将两个值求和后更新缓存。
        """
        import openpyxl
        wb = openpyxl.load_workbook(self.excel_path, data_only=True)

        for src_alias in ["当月附表2", "当月奉贤附表2", "上年同期附表2"]:
            cache = self.cache.get(src_alias, {})
            if not cache:
                continue

            cfg = self.sources.get(src_alias, {})
            sheet_name = cfg.get("sheet", "")
            label_col = cfg.get("label_col", 0)

            ws = self._match_sheet(wb, sheet_name)
            if ws is None:
                continue

            total = 0
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                label = self._norm(row[label_col].value)
                if label == '化验收入':
                    val = row[1].value  # col B (index 1) = 当月
                    if isinstance(val, (int, float)):
                        total += val

            if total > 0:
                row_data = cache.get('化验收入', {})
                old_val = row_data.get(1, 0)
                row_data[1] = total
                cache['化验收入'] = row_data
                print(f"  [REPAIR] {src_alias} 化验收入: {old_val:.2f} → {total:.2f}"
                      f" (门急诊+住院合计, diff={total - old_val:.2f})")

        wb.close()

    def get_value(self, source_alias, row_label, col_idx):
        """获取指定 source 中指定行标签、列的值，自动处理万元转换."""
        if source_alias not in self.cache:
            return None
        sheet_data = self.cache[source_alias]
        row = self._fuzzy_get(sheet_data, self._norm(row_label))
        if row is None:
            return None
        actual_col = col_idx + self.label_col.get(source_alias, 0)
        val = row.get(actual_col)
        if val is None:
            return None
        if not isinstance(val, (int, float)):
            return val
        if self.needs_wan.get(source_alias, False):
            val = val / 10000.0
        return val

    def get_raw_value(self, source_alias, row_label, col_idx):
        """Like get_value but without wan conversion. Used for sum computations."""
        if source_alias not in self.cache:
            return None
        sheet_data = self.cache[source_alias]
        row = self._fuzzy_get(sheet_data, self._norm(row_label))
        if row is None:
            return None
        actual_col = col_idx + self.label_col.get(source_alias, 0)
        val = row.get(actual_col)
        if val is None:
            return None
        if not isinstance(val, (int, float)):
            return val
        return val


# ---------------------------------------------------------------------------
# MaterialDetailLoader
# ---------------------------------------------------------------------------

class MaterialDetailLoader:
    """加载卫生材料明细 Excel，提供按条件聚合查询能力.

    明细文件有两个 sheet，列布局不同：
    - 2026 sheet (20列): col 7=院区, col 16=是否收费, col 12=入账科目, col 3=部门名称, col 19=金额
    - 2025 sheet (19列): 无院区列, col 15=是否收费, col 11=入账科目, col 3=部门名称, col 18=金额
    """

    def __init__(self, detail_path):
        self.detail_path = detail_path
        self.data = {}  # year -> list of dicts
        self._load()

    def _load(self):
        wb = openpyxl.load_workbook(self.detail_path, data_only=True)
        for ws in wb.worksheets:
            if ws.max_column < 2:
                continue
            ncols = ws.max_column
            if ncols == 20:
                year = 2026
                col_map = {"period": 1, "campus": 7, "chargeable": 16, "subject": 12, "dept": 3, "amount": 19}
            elif ncols == 19:
                year = 2025
                col_map = {"period": 1, "chargeable": 15, "subject": 11, "dept": 3, "amount": 18}
            else:
                continue

            rows = []
            for r in range(2, ws.max_row + 1):
                period_raw = str(ws.cell(r, col_map["period"] + 1).value or "").strip()
                chargeable_raw = str(ws.cell(r, col_map["chargeable"] + 1).value or "").strip()
                subject = str(ws.cell(r, col_map["subject"] + 1).value or "").strip()
                dept = str(ws.cell(r, col_map["dept"] + 1).value or "").strip()
                amount = ws.cell(r, col_map["amount"] + 1).value
                if not isinstance(amount, (int, float)):
                    continue
                row_data = {
                    "period": period_raw,
                    "chargeable": chargeable_raw == "是",
                    "subject": subject,
                    "dept": dept,
                    "amount": amount,
                }
                if "campus" in col_map:
                    row_data["campus"] = str(ws.cell(r, col_map["campus"] + 1).value or "").strip()
                rows.append(row_data)

            self.data[year] = rows
            campus_info = "with campus" if "campus" in col_map else "no campus"
            print(f"  [LOAD] 卫生材料明细 {year} ← {ws.title} ({len(rows)} rows, {campus_info})")
        wb.close()

    def aggregate(self, year, chargeable=None, subject_contains=None,
                  dept_contains=None, campus=None, period=None):
        """按条件过滤并返回金额总和（元）.

        Args:
            year: 2026 or 2025
            chargeable: True=是, False=否, None=不限
            subject_contains: 入账科目包含此字符串
            dept_contains: 部门名称包含此字符串
            campus: 院区（仅2026数据有效）
            period: 期间过滤，如 "2026-04"（None=不限）

        Returns:
            金额总和（元），若无匹配返回 0.
        """
        if year not in self.data:
            return 0
        total = 0.0
        for row in self.data[year]:
            if period and row["period"] != period:
                continue
            if chargeable is not None and row["chargeable"] != chargeable:
                continue
            if subject_contains and subject_contains not in row["subject"]:
                continue
            if dept_contains and dept_contains not in row["dept"]:
                continue
            if campus and row.get("campus") and campus not in row["campus"]:
                continue
            total += row["amount"]
        return total


# ---------------------------------------------------------------------------
# ComputedValueResolver
# ---------------------------------------------------------------------------

class ComputedValueResolver:
    """Resolves computed column and text variable specifications."""

    def __init__(self, loader, data_month=None, detail_loader=None):
        self.loader = loader
        self.data_month = data_month or 4
        self.detail_loader = detail_loader

    def resolve(self, compute_spec, resolved_cols=None):
        ctype = compute_spec["type"]
        if ctype == "value":
            return compute_spec["value"]
        elif ctype == "subtract":
            a = resolved_cols[compute_spec["a_col"]]
            b = resolved_cols[compute_spec["b_col"]]
            return a - b if a is not None and b is not None else None
        elif ctype == "add":
            a = resolved_cols[compute_spec["a_col"]]
            b = resolved_cols[compute_spec["b_col"]]
            return a + b if a is not None and b is not None else None
        elif ctype == "divide":
            a = resolved_cols[compute_spec["a_col"]]
            b = resolved_cols[compute_spec["b_col"]]
            if a is not None and b is not None and b != 0:
                return a / b
            return None
        elif ctype == "multiply":
            a = resolved_cols[compute_spec["a_col"]]
            b = resolved_cols[compute_spec["b_col"]]
            return a * b if a is not None and b is not None else None
        elif ctype == "sum":
            total = 0
            for part in compute_spec["parts"]:
                v = self.loader.get_raw_value(part["source"], part.get("xl_label", ""), part["xl_col"])
                if v is not None and isinstance(v, (int, float)):
                    total += v
                else:
                    return None
            if compute_spec.get("needs_wan", False):
                total /= 10000.0
            return total
        elif ctype == "time_node_diff":
            exec_rate = resolved_cols[compute_spec["exec_rate_col"]]
            if exec_rate is None:
                return None
            month = compute_spec.get("month", self.data_month)
            return exec_rate - (month / 12.0)
        elif ctype == "pct_diff":
            a = resolved_cols[compute_spec["a_col"]]
            b = resolved_cols[compute_spec["b_col"]]
            if a is not None and b is not None and b != 0:
                return (a - b) / abs(b)
            return None
        elif ctype == "pct_change":
            change = resolved_cols[compute_spec["change_col"]]
            base = resolved_cols[compute_spec["base_col"]]
            if change is not None and base is not None:
                denominator = base - change
                if denominator != 0:
                    return change / abs(denominator)
            return None
        elif ctype == "divide_excel":
            a = self._resolve_source(compute_spec["a"])
            b = self._resolve_source(compute_spec["b"])
            if a is not None and b is not None and b != 0:
                return a / b
            return None
        elif ctype == "time_node_diff_excel":
            exec_rate = None
            if "exec_rate" in compute_spec:
                exec_rate = self._resolve_source(compute_spec["exec_rate"])
            elif "a" in compute_spec and "b" in compute_spec:
                a = self._resolve_source(compute_spec["a"])
                b = self._resolve_source(compute_spec["b"])
                if a is not None and b is not None and b != 0:
                    exec_rate = a / b
            if exec_rate is None:
                return None
            month = compute_spec.get("month", self.data_month)
            return exec_rate - (month / 12.0)
        elif ctype == "subtract_excel":
            a = self._resolve_source(compute_spec["a"])
            b = self._resolve_source(compute_spec["b"])
            return a - b if a is not None and b is not None else None
        elif ctype == "sum_excel":
            total = 0
            for part in compute_spec["parts"]:
                # Use get_raw_value to avoid double-dividing by 10000
                v = self.loader.get_raw_value(part["source"], part.get("xl_label", ""), part["xl_col"])
                if v is not None and isinstance(v, (int, float)):
                    total += v
                else:
                    return None
            if compute_spec.get("needs_wan", False):
                total /= 10000.0
            return total
        elif ctype == "detail_sum":
            if self.detail_loader is None:
                return None
            year = compute_spec["year"]
            month = compute_spec.get("month", self.data_month)
            period = f"{year}-{month:02d}"
            raw = self.detail_loader.aggregate(
                year=year,
                chargeable=compute_spec.get("chargeable"),
                subject_contains=compute_spec.get("subject_contains"),
                dept_contains=compute_spec.get("dept_contains"),
                campus=compute_spec.get("campus"),
                period=period,
            )
            if compute_spec.get("needs_wan", False):
                raw = raw / 10000.0
            return raw
        elif ctype == "add_refs":
            a = self._resolve_source(compute_spec["a"])
            b = self._resolve_source(compute_spec["b"])
            return a + b if a is not None and b is not None else None
        elif ctype == "subtract_refs":
            a = self._resolve_source(compute_spec["a"])
            b = self._resolve_source(compute_spec["b"])
            return a - b if a is not None and b is not None else None
        return None

    def _resolve_source(self, source_spec):
        """Resolve a value from an Excel source spec or compute spec."""
        if isinstance(source_spec, dict):
            if "compute" in source_spec:
                return self.resolve(source_spec["compute"])
            return self.loader.get_value(
                source_spec.get("source", ""),
                source_spec.get("xl_label", ""),
                source_spec.get("xl_col", 0)
            )
        return source_spec


# ---------------------------------------------------------------------------
# TextUpdater
# ---------------------------------------------------------------------------

class TextUpdater:
    """Updates text paragraphs in non-table shapes (text boxes)."""

    @staticmethod
    def norm(s):
        if s is None:
            return ""
        return re.sub(r'[\s　\n\r]+', '', str(s).strip())

    @classmethod
    def find_shape_by_name(cls, slide, shape_name):
        target = cls.norm(shape_name)
        for shape in slide.shapes:
            if cls.norm(shape.name) == target:
                return shape
        return None

    @staticmethod
    def format_var(value, format_type):
        if value is None:
            return "N/A"
        if not isinstance(value, (int, float)):
            return str(value)
        if format_type == "wan_int":
            return f"{round(value / 10000.0):,}"
        elif format_type == "wan_1dp":
            return f"{value / 10000.0:,.1f}"
        elif format_type == "pct_1dp":
            return f"{value * 100:.1f}%"
        elif format_type == "pct_2dp":
            return f"{value * 100:.2f}%"
        elif format_type == "pp_1dp":
            return f"{abs(value) * 100:.1f}"
        elif format_type == "pct_abs_1dp":
            return f"{abs(value) * 100:.1f}%"
        elif format_type == "int":
            return f"{round(value):,}"
        else:
            if isinstance(value, float) and abs(value) < 10 and value != int(value):
                return f"{value:.2f}"
            return f"{round(value):,}"

    @classmethod
    def update_text(cls, slide, text_config, loader, resolver, data_month=None):
        shape_name = text_config["shape_name"]
        shape = cls.find_shape_by_name(slide, shape_name)
        if shape is None:
            return False

        var_values = {}
        for var_name, var_spec in text_config.get("vars", {}).items():
            if var_spec.get("type") == "value":
                var_values[var_name] = var_spec["value"]
            elif "compute" in var_spec:
                var_values[var_name] = resolver.resolve(var_spec["compute"])
            else:
                val = loader.get_value(
                    var_spec.get("source", ""),
                    var_spec.get("xl_label", ""),
                    var_spec.get("xl_col", 0)
                )
                var_values[var_name] = val

        format_overrides = text_config.get("format_overrides", {})
        formatted_values = {}
        for var_name, val in var_values.items():
            fmt = format_overrides.get(var_name)
            if fmt:
                formatted_values[var_name] = cls.format_var(val, fmt)
            else:
                if isinstance(val, (int, float)):
                    formatted_values[var_name] = f"{round(val):,}"
                else:
                    formatted_values[var_name] = str(val) if val is not None else "N/A"

        template = text_config["template"]
        # Resolve {month} and {month:02d} in template text
        if data_month is not None:
            template = template.replace("{month:02d}", f"{data_month:02d}")
            template = template.replace("{month}", str(data_month))
        new_text = template.format(**formatted_values)

        if shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = new_text
            # Copy font properties if original had them
            return True

        return False


# ---------------------------------------------------------------------------
# PPTTableFinder
# ---------------------------------------------------------------------------

class PPTTableFinder:

    @staticmethod
    def norm(s):
        if s is None:
            return ""
        return re.sub(r'[\s　\n\r]+', '', str(s).strip())

    @classmethod
    def find_table_by_header(cls, slide, headers, header_row=0):
        """在 slide 中查找包含指定表头的表格."""
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            table = shape.table
            if header_row >= len(table.rows):
                continue
            row_cells = [cls.norm(table.rows[header_row].cells[c].text) for c in range(len(table.columns))]
            # Check if all headers appear as substrings in the row
            match_count = 0
            for h in headers:
                h_norm = cls.norm(h)
                for cell_text in row_cells:
                    if h_norm and h_norm in cell_text:
                        match_count += 1
                        break
            if match_count >= len(headers) * 0.6:  # 60% threshold
                return table
        return None

    @classmethod
    def find_row_by_label(cls, table, label, col_idx=0):
        """在表格第 col_idx 列中查找包含 label 的行号."""
        target = cls.norm(label)
        for r_idx, row in enumerate(table.rows):
            cell_text = cls.norm(row.cells[col_idx].text)
            if target and target in cell_text:
                return r_idx
        # Try looser match
        for r_idx, row in enumerate(table.rows):
            cell_text = cls.norm(row.cells[col_idx].text)
            if target and cell_text and cell_text in target:
                return r_idx
        return None

    @classmethod
    def get_row_labels(cls, table, col_idx=0):
        """获取表格指定列的所有行标签."""
        return [cls.norm(table.rows[r].cells[col_idx].text) for r in range(len(table.rows))]


# ---------------------------------------------------------------------------
# ValueFormatter
# ---------------------------------------------------------------------------

class ValueFormatter:
    """格式化更新到 PPT 单元格的值."""

    @staticmethod
    def is_percentage_text(val):
        """判断原始文本是否包含百分号."""
        if isinstance(val, str):
            return '%' in val
        return False

    @staticmethod
    def is_decimal(val):
        """判断值是否为小数（非金额）."""
        if isinstance(val, float) and abs(val) < 100 and val != int(val):
            return True
        return False

    @classmethod
    def format_value(cls, raw_value, ppt_original_text, format_override=None):
        """根据 PPT 原始格式和新值，决定输出文本.

        Args:
            raw_value: 从 Excel 获取的新值（数值或字符串）.
            ppt_original_text: PPT 单元格原始文本，用于判断格式.
            format_override: 可选的格式字符串，如 "2dp", "pct", "int".

        Returns:
            格式化后的字符串.
        """
        if raw_value is None:
            return ppt_original_text  # 保持原样

        # 如果新值是字符串（如百分比），直接返回
        if isinstance(raw_value, str):
            return raw_value

        if not isinstance(raw_value, (int, float)):
            return str(raw_value)

        # Explicit format override takes priority
        if format_override == "2dp":
            return f"{raw_value:,.2f}"
        elif format_override == "int":
            return f"{round(raw_value):,}"
        elif format_override == "pct":
            return f"{raw_value * 100:.1f}%"

        # 判断原始 PPT 单元格的格式
        ppt_text = str(ppt_original_text).strip() if ppt_original_text else ""

        # 包含 % → 按百分比格式化（Excel 百分比数据均为比值形式：0.3459→34.6%, 1→100%, 1.1875→118.8%）
        if '%' in ppt_text:
            return f"{raw_value * 100:.1f}%"

        # 负值 → 带符号
        val = raw_value

        # 根据原始文本判断是否需要小数
        if re.search(r'\.\d', ppt_text):
            # 原始有小数 → 保留小数
            return f"{val:,.2f}"
        else:
            # 原始为整数 → 四舍五入
            rounded = round(val)
            return f"{rounded:,}"


# ---------------------------------------------------------------------------
# CostAnalysisUpdater
# ---------------------------------------------------------------------------

class CostAnalysisUpdater:
    """主控制器：根据 mapping 更新 PPT."""

    def __init__(self, excel_path, pptx_path, mapping_path):
        self.excel_path = excel_path
        self.pptx_path = pptx_path
        self.mapping_path = mapping_path

        with open(mapping_path, 'r', encoding='utf-8') as f:
            self.mapping = json.load(f)

        self.loader = ExcelDataLoader(excel_path, self.mapping.get("excel_sources", {}))
        self.prs = Presentation(pptx_path)
        self.finder = PPTTableFinder()
        self.formatter = ValueFormatter()

        self.data_month = detect_data_month(excel_path)

        # Auto-detect 卫生材料明细 file in same directory as main Excel
        self.detail_loader = None
        excel_dir = os.path.dirname(os.path.abspath(excel_path))
        for fname in os.listdir(excel_dir):
            if '卫生材料明细' in fname and fname.endswith('.xlsx') and not fname.startswith('~$'):
                detail_path = os.path.join(excel_dir, fname)
                print(f"  [DETAIL] Found: {fname}")
                self.detail_loader = MaterialDetailLoader(detail_path)
                break
        if self.detail_loader is None:
            print("  [DETAIL] No detail file found, detail_sum compute type will return None")

        self.resolver = ComputedValueResolver(self.loader, self.data_month, self.detail_loader)
        self.text_updater = TextUpdater()

        self.stats = {"updated_cells": 0, "skipped_rows": 0, "skipped_cells": 0, "slides_processed": 0}

    def run(self, output_path):
        slides_config = self.mapping.get("slides", {})
        total_slides = len(slides_config)

        for slide_key, slide_cfg in slides_config.items():
            slide_num = int(slide_key)
            if slide_num < 1 or slide_num > len(self.prs.slides):
                print(f"  [SKIP] Slide {slide_num}: out of range")
                continue

            slide = self.prs.slides[slide_num - 1]  # 1-indexed → 0-indexed
            print(f"\n--- Slide {slide_num}: {slide_cfg.get('desc', '')} ---")
            self._process_slide(slide, slide_cfg)

        self.prs.save(output_path)
        print(f"\n{'='*50}")
        print(f"Summary: Updated {self.stats['updated_cells']} cells across "
              f"{self.stats['slides_processed']} slides, "
              f"skipped {self.stats['skipped_rows']} rows, "
              f"{self.stats['skipped_cells']} cells.")
        print(f"Output: {output_path}")

    def _resolve_headers(self, headers):
        """Replace {month}, {month:02d} etc. in header strings with actual data month."""
        resolved = []
        for h in headers:
            h = h.replace("{month:02d}", f"{self.data_month:02d}")
            h = h.replace("{month}", str(self.data_month))
            resolved.append(h)
        return resolved

    def _process_slide(self, slide, slide_cfg):
        headers = self._resolve_headers(slide_cfg.get("headers", []))
        header_row = slide_cfg.get("header_row", 0)

        # Phase 1: Table updates
        table = self.finder.find_table_by_header(slide, headers, header_row)
        if table is None:
            print(f"  [SKIP] Table not found with headers: {headers}")
        else:
            self.stats["slides_processed"] += 1
            rows_cfg = slide_cfg.get("rows", [])
            for row_cfg in rows_cfg:
                self._process_table_row(table, row_cfg, slide_cfg)

        # Phase 2: Text paragraph updates (independent of table existence)
        texts_cfg = slide_cfg.get("texts", [])
        for text_cfg in texts_cfg:
            success = self.text_updater.update_text(slide, text_cfg, self.loader, self.resolver, self.data_month)
            if success:
                print(f"  [UPD] Text shape '{text_cfg['shape_name']}' updated")
            else:
                print(f"  [SKIP] Text shape '{text_cfg['shape_name']}' not found")

    def _process_table_row(self, table, row_cfg, slide_cfg):
        label = row_cfg["label"]
        default_source = row_cfg.get("source")
        default_xl_label = row_cfg.get("xl_label", label)

        row_idx = self.finder.find_row_by_label(table, label, col_idx=0)
        if row_idx is None:
            print(f"  [SKIP Row] '{label}' not found in table")
            self.stats["skipped_rows"] += 1
            return

        col_entries = row_cfg.get("cols", [])
        resolved = {}   # {ppt_col: value}
        col_fmts = {}   # {ppt_col: format_override}

        # Pass 1: Resolve all simple (non-computed) column values
        for entry in col_entries:
            if isinstance(entry, list):
                ppt_col, xl_col = entry[0], entry[1]
                val = self.loader.get_value(default_source, default_xl_label, xl_col)
                resolved[ppt_col] = val
            elif isinstance(entry, dict):
                if "compute" in entry:
                    continue  # Defer to pass 2
                else:
                    ppt_col = entry["ppt"]
                    src = entry.get("source", default_source)
                    xlbl = entry.get("xl_label", default_xl_label)
                    xl_col = entry.get("xl", 0)
                    val = self.loader.get_value(src, xlbl, xl_col)
                    resolved[ppt_col] = val
                    if "format" in entry:
                        col_fmts[ppt_col] = entry["format"]

        # Pass 2: Evaluate computed columns (in order, so dependencies must be earlier)
        for entry in col_entries:
            if isinstance(entry, dict) and "compute" in entry:
                ppt_col = entry["ppt"]
                val = self.resolver.resolve(entry["compute"], resolved_cols=resolved)
                resolved[ppt_col] = val
                if "format" in entry:
                    col_fmts[ppt_col] = entry["format"]

        # Apply all resolved values to table cells
        for ppt_col, new_val in resolved.items():
            if new_val is None:
                print(f"  [SKIP Cell] Row='{label}', PPT col={ppt_col}: no value resolved")
                self.stats["skipped_cells"] += 1
                continue

            cell = table.rows[row_idx].cells[ppt_col]
            original_text = cell.text
            fmt = col_fmts.get(ppt_col)
            formatted = self.formatter.format_value(new_val, original_text, format_override=fmt)

            if formatted != original_text:
                self._update_cell_text(cell, formatted)
                self.stats["updated_cells"] += 1
                print(f"  [UPD] [{row_idx},{ppt_col}] '{label}': \"{original_text[:20]}\" → \"{formatted}\"")

    @staticmethod
    def _update_cell_text(cell, new_text):
        """Write new text to a table cell, replacing existing content."""
        if cell.text_frame.paragraphs:
            for para in cell.text_frame.paragraphs:
                if para.runs:
                    para.runs[0].text = new_text
                else:
                    para.add_run().text = new_text
                break
        else:
            cell.text_frame.paragraphs[0].add_run().text = new_text


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="新华医院成本分析 PPT 数据更新工具")
    parser.add_argument("excel", help="Excel 底稿文件路径")
    parser.add_argument("pptx", help="PPT 模板文件路径")
    parser.add_argument("--output", "-o", default=None, help="输出路径（默认在原 PPT 名加日期后缀）")
    parser.add_argument("--mapping", "-m", default=None, help="mapping.json 路径")
    args = parser.parse_args()

    # 验证输入文件
    for path, desc in [(args.excel, "Excel"), (args.pptx, "PPTX")]:
        if not os.path.exists(path):
            print(f"[ERROR] {desc} file not found: {path}")
            sys.exit(1)

    # 默认 mapping 路径
    if args.mapping is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        args.mapping = os.path.join(script_dir, "..", "references", "mapping.json")

    if not os.path.exists(args.mapping):
        print(f"[ERROR] Mapping file not found: {args.mapping}")
        sys.exit(1)

    # 默认输出路径
    if args.output is None:
        base, ext = os.path.splitext(args.pptx)
        date_str = datetime.now().strftime("%m%d")
        args.output = f"{base}_updated_{date_str}{ext}"

    print(f"Excel:  {args.excel}")
    print(f"PPTX:   {args.pptx}")
    print(f"Output: {args.output}")
    print(f"Mapping: {args.mapping}")
    print()

    updater = CostAnalysisUpdater(args.excel, args.pptx, args.mapping)
    updater.run(args.output)


if __name__ == "__main__":
    main()
