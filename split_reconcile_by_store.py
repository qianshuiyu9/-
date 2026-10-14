# -*- coding: utf-8 -*-
"""
供应商对账单 - 按门店拆分 + 门店金额汇总
=========================================
读取「供应商账单」目录下的全部对账文件（每文件一个供应商），按门店拆分写入「门店对账」目录：
  - 一个门店一个文件：文件名 = 编号 + 号店.xlsx
  - 分公司 / 总部其他 聚合文件（多店合并）
  - 19 / 30 号店（油菜花/娱乐至上平台数据）单独出表：19号店.xlsx / 30号店.xlsx，
    不与 总部其他.xlsx 合并；明细从源文件尾部「采购单/所属门店」明细块解析
  - 不开票供应商（配置 NO_INVOICE_SUPPLIERS）：其所有门店数据一律并入 总部其他.xlsx，不按门店拆分
  - 文件内一个供应商一个 Sheet；同名 Sheet 已存在则跳过，绝不覆盖 / 删除已有 Sheet
  - 每个输出文件自动从旧「按店铺对账」同名文件复制「汇总」Sheet（收款信息/备注等原样保留，
    跨表公式改写为静态数值），并按本期数据更新金额；新供应商自动加行（收款信息留空待补）；
    已存在的「汇总」仅在金额与本期不一致时刷新（其余内容不动）
  - 「汇总」列宽/行高自适应：按内容自动调整列宽并启用自动换行，保证收款信息、金额完整显示
  - 最后生成 门店金额汇总.xlsx

金额口径（与旧「按店铺对账」文件逐项核对一致，相乘不四舍五入）：
  下单合计 = Σ下单金额（含平台行金额：付款说明为数值时优先取之）
  减款合计 = Σ调整额（误差备注右侧数值列，平台行的优惠备注不额外扣减）
  折后合计 = 下单合计 + 减款合计
  不开票金额 = 折后合计 × 折扣系数
  开票金额   = 不开票金额 × 开票系数（加点=1.01，不加点=1.0）
  应付金额   = 分公司/45/48/总部其他 → 不开票金额；单独门店 → 开票金额
  汇总Sheet：单店/分公司/总部其他 = 应付/开票口径；45/48 = 不开票口径

运行：
  python split_reconcile_by_store.py
  （可选：python split_reconcile_by_store.py --input 目录 --output 目录）
"""

import argparse
import re
import sys
from copy import copy
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import supplier_config

# ============================================================
# 配置区（门店归类 / 供应商规则，均为待替换的默认初值）
# ============================================================
INPUT_DIR  = r"C:\Users\Administrator\Desktop\供应商账单"
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\门店对账"
OLD_DIR    = r"C:\Users\Administrator\Desktop\按店铺对账"   # 旧文件目录：汇总Sheet模板（收款信息）来源
SUMMARY_NAME = "门店金额汇总.xlsx"

BRANCH_FILE   = "分公司.xlsx"
HQ_OTHER_FILE = "总部其他.xlsx"

# 门店归属 + 供应商规则（开票状态/折扣/收款信息）已迁移至 supplier_config.py 的 supplier_config.json，
# 由 GUI 界面维护，脚本运行时通过 supplier_config 模块动态读取。

# 油菜花（娱乐至上）平台数据：源文件尾部明细块「所属门店」关键词 → 门店编号
PLATFORM_STORE_MAP = {"红光沃尔玛": "30", "锦江大融城": "19"}

# 分组标签（决定应付金额取值）
GROUP_BRANCH = "分公司"
GROUP_HQ     = "总部其他"
GROUP_SINGLE = "单独门店"
NO_INVOICE_GROUPS = (GROUP_BRANCH, GROUP_HQ)   # 这些分组不开票：应付金额 = 不开票金额

DETAIL_COLUMNS = ["单号", "供应商", "收货门店", "种数", "个数", "收货数量", "收货进度",
                  "下单金额", "收货价值", "付款进度", "额外费用", "下单时间", "采购人",
                  "误差备注", "调整额"]
COL_AMOUNT = DETAIL_COLUMNS.index("下单金额") + 1   # 8 → H
COL_ADJ    = DETAIL_COLUMNS.index("调整额") + 1     # 15 → O

STORE_RE = re.compile(r"^(\d+)号店")
ILLEGAL_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")

# ============================================================
# 样式（沿用旧脚本风格）
# ============================================================
THIN_BORDER = Border(
    left=Side(style="thin", color="D9DEE7"),
    right=Side(style="thin", color="D9DEE7"),
    top=Side(style="thin", color="D9DEE7"),
    bottom=Side(style="thin", color="D9DEE7"),
)
HEADER_FILL = PatternFill(fill_type="solid", fgColor="EAF2FF")
HEADER_FONT = Font(bold=True, color="1F2937")
ZEBRA_FILL_1 = PatternFill(fill_type="solid", fgColor="FFFFFF")
ZEBRA_FILL_2 = PatternFill(fill_type="solid", fgColor="F7F9FC")
TOTAL_FILL = PatternFill(fill_type="solid", fgColor="FFF3CD")
TOTAL_FONT = Font(bold=True, color="1F2937")


# ============================================================
# 工具函数
# ============================================================
def round2(v):
    """half-up 四舍五入保留 2 位"""
    return float(Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def norm_store(v):
    """门店名规范化：去掉全部空白/换行"""
    if v is None:
        return ""
    return re.sub(r"\s+", "", str(v))


def parse_month(v):
    """下单时间 → 'YYYY-MM'"""
    if isinstance(v, datetime):
        return f"{v.year:04d}-{v.month:02d}"
    m = re.search(r"(\d{4})[-/](\d{1,2})", str(v))
    return f"{m.group(1)}-{int(m.group(2)):02d}" if m else None


def period_str(months):
    """账期：单月 '2026-09'，跨月 '2026-08~2026-09'"""
    ms = sorted(set(m for m in months if m))
    if not ms:
        return ""
    return ms[0] if len(ms) == 1 else f"{ms[0]}~{ms[-1]}"


def clean_sheet_name(name):
    """Sheet 名清洗非法字符并截断 31 字符"""
    s = ILLEGAL_SHEET_CHARS.sub("", str(name)).strip()
    return s[:31] or "Sheet"


def supplier_rule(name):
    """折扣/开票系数：从供应商配置读取（GUI 维护）"""
    return supplier_config.get_rule(name)


def discount_label(d):
    """折扣标签：0.99 → '99折'"""
    if abs(d - 1.0) < 1e-9:
        return None
    v = Decimal(str(d)) * 100
    if v == v.to_integral_value():
        return f"{int(v)}折"
    return f"折扣×{d}"


def route(code, supplier=None):
    """门店编号 + 供应商 → (输出文件名, 分组, 是否未归类)
    门店归属三档（GUI 维护，supplier_config.stores）：单独拆分 / 总部其他 / 分公司；
    「单独拆分」再按 invoice 口径分：开票=GROUP_SINGLE，不开票=GROUP_BRANCH。
    不开票供应商：仅「单独拆分(不开票口径)」「分公司」两类门店拆出，其余并入 总部其他.xlsx"""
    code = str(code).strip()
    store = supplier_config.get_store(code)

    if supplier is not None and supplier_config.is_no_invoice(supplier):
        if store is not None:
            g = store.get("group")
            if g == supplier_config.STORE_GROUP_BRANCH:
                return (BRANCH_FILE, GROUP_BRANCH, False)
            if (g == supplier_config.STORE_GROUP_SPLIT
                    and store.get("invoice") == supplier_config.INVOICE_OFF):
                return (f"{code}号店.xlsx", GROUP_BRANCH, False)
        return (HQ_OTHER_FILE, GROUP_HQ, False)

    if store is None:
        return (HQ_OTHER_FILE, GROUP_HQ, True)
    g = store.get("group")
    if g == supplier_config.STORE_GROUP_BRANCH:
        return (BRANCH_FILE, GROUP_BRANCH, False)
    if g == supplier_config.STORE_GROUP_HQ:
        return (HQ_OTHER_FILE, GROUP_HQ, False)
    if g == supplier_config.STORE_GROUP_SPLIT:
        grp = GROUP_SINGLE if store.get("invoice") == supplier_config.INVOICE_ON else GROUP_BRANCH
        return (f"{code}号店.xlsx", grp, False)
    return (HQ_OTHER_FILE, GROUP_HQ, True)


def calc_amounts(rows, rule):
    """按行集合计算金额链"""
    total = sum(r["amount"] for r in rows)
    adj = sum(r["adj"] for r in rows)
    after_adj = total + adj
    after_disc = round2(after_adj * rule["discount"])
    no_inv = after_disc
    inv = round2(after_disc * rule["invoice"])
    return {
        "total": round2(total),
        "adj": round2(adj),
        "after_adj": round2(after_adj),
        "after_disc": after_disc,
        "no_invoice": no_inv,
        "invoice": inv,
    }


def calc_unrounded(rows, rule):
    """按行集合计算不四舍五入的 (不开票金额, 开票金额)（用于汇总Sheet，与旧文件数值精确一致）"""
    net = sum((Decimal(str(r["amount"])) + Decimal(str(r["adj"])) for r in rows),
              Decimal("0"))
    no_inv = net * Decimal(str(rule["discount"]))
    inv = no_inv * Decimal(str(rule["invoice"]))
    return float(no_inv), float(inv)


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ============================================================
# 读取源表
# ============================================================
def _read_rows(path):
    """按后缀读取首个 sheet 的全部行 → list[list]（对齐 openpyxl iter_rows 语义）"""
    if path.suffix.lower() == ".xls":
        return _read_xls_rows(path)
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    all_rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return all_rows


def _read_xls_rows(path):
    """xlrd 读取 .xls 全部行，空单元格转 None、日期转 datetime"""
    import xlrd
    book = xlrd.open_workbook(str(path))
    sheet = book.sheet_by_index(0)
    rows = []
    for i in range(sheet.nrows):
        row = []
        for j in range(sheet.ncols):
            cell = sheet.cell(i, j)
            if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                row.append(None)
            elif cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    row.append(xlrd.xldate_as_datetime(cell.value, book.datemode))
                except Exception:
                    row.append(cell.value)
            else:
                row.append(cell.value)
        rows.append(row)
    return rows


def read_reconcile_rows(path):
    """读取一个源文件的有效对账行（门店匹配 'N号店' 且下单金额为数值）"""
    all_rows = _read_rows(path)
    hdr_i, col = None, {}
    for i, r in enumerate(all_rows[:20]):
        vals = [str(v).strip() if v is not None else "" for v in r]
        if "收货门店" in vals and "下单金额" in vals:
            hdr_i = i
            col = {v: j for j, v in enumerate(vals) if v}
            break
    if hdr_i is None:
        print(f"  [WARN] 未找到表头（收货门店/下单金额）：{path.name}，跳过")
        return []

    si = col["收货门店"]
    ai = col["下单金额"]
    mi = col.get("供应商", 1)
    di = col.get("下单时间")
    oi = col.get("误差备注", -1) + 1   # 调整额列 = 误差备注右侧一列

    recs = []
    for r in all_rows[hdr_i + 1:]:
        store = norm_store(r[si] if si < len(r) else None)
        m = STORE_RE.match(store)
        amt = r[ai] if ai < len(r) else None
        if not m or not isinstance(amt, (int, float)) or isinstance(amt, bool):
            continue
        ov = r[oi] if 0 <= oi < len(r) else None
        adj = float(ov) if isinstance(ov, (int, float)) and not isinstance(ov, bool) else 0.0
        if ov is not None and not isinstance(ov, (int, float)):
            try:
                adj = float(str(ov).strip()) if str(ov).strip() else 0.0
            except ValueError:
                adj = 0.0

        cells = {}
        for name in DETAIL_COLUMNS:
            j = col.get(name)
            cells[name] = r[j] if j is not None and j < len(r) else None
        if cells.get("单号") is not None:
            cells["单号"] = str(cells["单号"]).strip()
        cells["收货门店"] = store
        cells["供应商"] = str(r[mi]).strip() if mi is not None and mi < len(r) and r[mi] is not None else "未知供应商"
        cells["调整额"] = adj if adj else None

        recs.append({
            "code": m.group(1),
            "supplier": cells["供应商"],
            "amount": float(amt),
            "adj": adj,
            "month": parse_month(r[di] if di is not None and di < len(r) else None),
            "cells": cells,
        })
    return recs


def read_platform_rows(path):
    """读取源文件尾部油菜花（娱乐至上）平台明细块（表头含「采购单」+「所属门店」）。
    金额：付款说明(K)为数值时优先取之，否则取金额(D)；优惠(备注M)不额外扣减（与旧文件一致）。"""
    all_rows = _read_rows(path)
    hdr_i, col = None, {}
    for i, r in enumerate(all_rows):
        vals = [str(v).strip() if v is not None else "" for v in r]
        if "采购单" in vals and "所属门店" in vals:
            hdr_i = i
            col = {v: j for j, v in enumerate(vals) if v}
            break
    if hdr_i is None:
        return []

    def cell(r, name):
        j = col.get(name)
        return r[j] if j is not None and j < len(r) else None

    recs = []
    for r in all_rows[hdr_i + 1:]:
        store_raw = norm_store(cell(r, "所属门店"))
        if not store_raw:
            continue
        code = None
        for kw, c in PLATFORM_STORE_MAP.items():
            if kw in store_raw:
                code = c
                break
        if code is None:
            print(f"  [WARN] 平台行门店未在 PLATFORM_STORE_MAP 中配置，跳过：{path.name} / {store_raw}")
            continue

        k = cell(r, "付款说明")
        d = cell(r, "金额")
        amount = float(k) if is_number(k) else (float(d) if is_number(d) else 0.0)
        qty = cell(r, "数量")
        supplier = cell(r, "供应商")
        supplier = str(supplier).strip() if supplier is not None else "未知供应商"
        danhao = cell(r, "采购单")

        cells = {
            "单号": str(danhao).strip() if danhao is not None else None,
            "供应商": supplier,
            "收货门店": f"{code}号店-{store_raw}",
            "种数": None,
            "个数": qty,
            "收货数量": qty,
            "收货进度": cell(r, "状态"),
            "下单金额": amount,
            "收货价值": cell(r, "含税金额"),
            "付款进度": cell(r, "付款状态"),
            "额外费用": None,
            "下单时间": cell(r, "创建时间"),
            "采购人": cell(r, "操作人"),
            "误差备注": cell(r, "备注"),
            "调整额": None,
        }
        recs.append({
            "code": code,
            "supplier": supplier,
            "amount": amount,
            "adj": 0.0,
            "month": parse_month(cell(r, "创建时间")),
            "cells": cells,
        })
    return recs


# ============================================================
# 写入 Sheet
# ============================================================
def autofit(ws, ncols, max_rows=200):
    for j in range(1, ncols + 1):
        width = 0
        for i in range(1, min(ws.max_row, max_rows) + 1):
            v = ws.cell(row=i, column=j).value
            if v is None or (isinstance(v, str) and v.startswith("=")):
                continue
            width = max(width, sum(2 if ord(ch) > 127 else 1 for ch in str(v)))
        ws.column_dimensions[get_column_letter(j)].width = min(max(width + 2, 8), 50)


def write_supplier_sheet(wb, sheet_name, rows, group, period, rule):
    """新建 Sheet 并写入明细 + 汇总块（调用前须确保 Sheet 不存在）"""
    ws = wb.create_sheet(sheet_name)

    # 表头
    for j, name in enumerate(DETAIL_COLUMNS, start=1):
        c = ws.cell(row=1, column=j, value=name)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    # 明细行
    for i, rec in enumerate(rows, start=2):
        for j, name in enumerate(DETAIL_COLUMNS, start=1):
            c = ws.cell(row=i, column=j, value=rec["cells"].get(name))
            c.border = THIN_BORDER
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.fill = ZEBRA_FILL_2 if i % 2 == 1 else ZEBRA_FILL_1
        ws.cell(row=i, column=1).number_format = "@"                 # 单号=文本
        ws.cell(row=i, column=COL_AMOUNT).number_format = "0.00"     # 下单金额
        ws.cell(row=i, column=9).number_format = "0.00"              # 收货价值
        ws.cell(row=i, column=COL_ADJ).number_format = "0.00"        # 调整额
        ws.cell(row=i, column=12).number_format = "yyyy-mm-dd hh:mm"

    d = 1 + len(rows)                       # 最后一行明细行号
    amt_letter = get_column_letter(COL_AMOUNT)
    adj_letter = get_column_letter(COL_ADJ)
    amounts = calc_amounts(rows, rule)

    block = []                              # [(标签, 金额公式或值)]（行号 = 明细末行 + 序号）
    block.append(("合计", f"=SUM({amt_letter}2:{amt_letter}{d})"))
    if abs(amounts["adj"]) >= 0.005:
        block.append(("减款合计", f"=SUM({adj_letter}2:{adj_letter}{d})"))

    r = d + 1                               # 合计行号
    ref_adj = f"{amt_letter}{r + 1}" if len(block) == 2 else None
    if ref_adj:
        block.append(("折后合计", f"={amt_letter}{r}+{ref_adj}"))
    else:
        block.append(("折后合计", f"={amt_letter}{r}"))
    r_disc = d + len(block)                 # 折后合计行号

    d_label = discount_label(rule["discount"])
    if d_label:
        block.append((d_label, f"={amt_letter}{r_disc}*{rule['discount']}"))
        r_base = d + len(block)
    else:
        r_base = r_disc

    block.append(("不开票金额", f"={amt_letter}{r_base}"))
    r_no_inv = d + len(block)

    if abs(rule["invoice"] - 1.0) > 1e-9:
        inv_label = "开票金额（加一个点）"
        inv_formula = f"={amt_letter}{r_no_inv}*{rule['invoice']}"
    else:
        inv_label = "开票金额（不加点）"
        inv_formula = f"={amt_letter}{r_no_inv}"
    block.append((inv_label, inv_formula))
    r_inv = d + len(block)

    if group in NO_INVOICE_GROUPS:
        pay_label, pay_ref = "应付金额（不开票）", r_no_inv
    else:
        pay_label, pay_ref = "应付金额", r_inv
    block.append((pay_label, f"={amt_letter}{pay_ref}"))
    r_pay = d + len(block)

    # 写入汇总块（行号 = 明细末行 + 块内序号）
    for idx, (label, value) in enumerate(block):
        row_no = d + 1 + idx
        for j in range(1, len(DETAIL_COLUMNS) + 1):
            c = ws.cell(row=row_no, column=j)
            c.fill = TOTAL_FILL
            c.border = THIN_BORDER
        ca = ws.cell(row=row_no, column=1, value=label)
        ca.font = TOTAL_FONT
        ca.alignment = Alignment(horizontal="left", vertical="center")
        cv = ws.cell(row=row_no, column=COL_AMOUNT, value=value)
        cv.font = TOTAL_FONT
        cv.number_format = "0.00"
        cv.alignment = Alignment(horizontal="center", vertical="center")
    if period:
        cp = ws.cell(row=r_pay, column=2, value=f"账期：{period}")
        cp.font = TOTAL_FONT
        cp.alignment = Alignment(horizontal="left", vertical="center")

    autofit(ws, len(DETAIL_COLUMNS))


# ============================================================
# 门店金额汇总.xlsx
# ============================================================
def output_label(file_name, codes):
    cs = sorted(set(codes), key=int)
    if file_name == BRANCH_FILE:
        return "分公司（" + "、".join(cs) + "号店）"
    if file_name == HQ_OTHER_FILE:
        return "总部其他（" + "、".join(cs) + "号店）"
    if len(cs) == 1:
        return f"{cs[0]}号店"
    return f"{Path(file_name).stem}（" + "、".join(cs) + "号店）"


def write_summary(summary_path, file_order, files_map, supplier_order):
    """汇总矩阵：行=输出文件/分组；列=门店|账期|每供应商(开票/不开票)|合计(开票/不开票)"""
    wb = Workbook()
    ws = wb.active
    ws.title = "汇总"

    # 表头两行
    header = [(("门店", "门店")), (("账期", "账期"))]
    col = 3
    sup_cols = {}
    for sup in supplier_order:
        sup_cols[sup] = col
        ws.cell(row=1, column=col, value=sup)
        ws.cell(row=2, column=col, value="开票金额")
        ws.cell(row=2, column=col + 1, value="不开票金额")
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        col += 2
    col_total = col
    ws.cell(row=1, column=col_total, value="合计")
    ws.cell(row=2, column=col_total, value="开票金额")
    ws.cell(row=2, column=col_total + 1, value="不开票金额")
    ws.merge_cells(start_row=1, start_column=col_total, end_row=1, end_column=col_total + 1)
    ncols = col_total + 1

    ws.cell(row=1, column=1, value="门店")
    ws.cell(row=2, column=1, value="门店")
    ws.cell(row=1, column=2, value="账期")
    ws.cell(row=2, column=2, value="账期")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)

    for j in range(1, ncols + 1):
        for i in (1, 2):
            c = ws.cell(row=i, column=j)
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
            c.border = THIN_BORDER
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 数据行
    r = 3
    tot_inv, tot_no_inv = 0.0, 0.0
    for fn in file_order:
        items = files_map[fn]
        codes = [c for _, it in items for c in it["codes"]]
        months = [m for _, it in items for m in it["months"]]
        ws.cell(row=r, column=1, value=output_label(fn, codes))
        ws.cell(row=r, column=2, value=period_str(months))
        row_inv, row_no_inv = 0.0, 0.0
        for sup, it in items:
            amounts = calc_amounts(it["rows"], supplier_rule(sup))
            c1 = sup_cols[sup]
            ws.cell(row=r, column=c1, value=amounts["invoice"]).number_format = "0.00"
            ws.cell(row=r, column=c1 + 1, value=amounts["no_invoice"]).number_format = "0.00"
            row_inv += amounts["invoice"]
            row_no_inv += amounts["no_invoice"]
        ws.cell(row=r, column=col_total, value=round2(row_inv)).number_format = "0.00"
        ws.cell(row=r, column=col_total + 1, value=round2(row_no_inv)).number_format = "0.00"
        tot_inv += row_inv
        tot_no_inv += row_no_inv
        for j in range(1, ncols + 1):
            c = ws.cell(row=r, column=j)
            c.border = THIN_BORDER
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.fill = ZEBRA_FILL_2 if r % 2 == 1 else ZEBRA_FILL_1
        r += 1

    # 总计行
    ws.cell(row=r, column=1, value="总计")
    ws.cell(row=r, column=col_total, value=round2(tot_inv)).number_format = "0.00"
    ws.cell(row=r, column=col_total + 1, value=round2(tot_no_inv)).number_format = "0.00"
    for j in range(1, ncols + 1):
        c = ws.cell(row=r, column=j)
        c.fill = TOTAL_FILL
        c.border = THIN_BORDER
        c.font = TOTAL_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "B3"
    autofit(ws, ncols)
    wb.save(summary_path)


# ============================================================
# 「汇总」Sheet：从旧文件复制 + 更新本期金额
# ============================================================
def old_summary_source(out_name):
    """输出文件名 → 旧「按店铺对账」同名文件路径（'20号店.xlsx' → '20.xlsx'）"""
    stem = Path(out_name).stem
    if stem.endswith("号店"):
        stem = stem[:-2]
    p = Path(OLD_DIR) / f"{stem}.xlsx"
    return p if p.exists() else None


def copy_summary_sheet(src_path, wb, title="汇总"):
    """把旧文件的「汇总」Sheet 复制为新文件的第一个 Sheet（样式/列宽/行高/合并原样保留）。
    公式单元格（跨表引用、总计）改写为静态数值（取缓存值占位），返回 (ws, 公式单元格集合)"""
    wb_src = load_workbook(src_path)
    ws_src = wb_src["汇总"]
    wb_val = load_workbook(src_path, data_only=True)
    ws_val = wb_val["汇总"]

    ws = wb.create_sheet(title, 0)
    for key, dim in ws_src.column_dimensions.items():
        ws.column_dimensions[key].width = dim.width
        ws.column_dimensions[key].hidden = dim.hidden
    for key, dim in ws_src.row_dimensions.items():
        ws.row_dimensions[key].height = dim.height
        ws.row_dimensions[key].hidden = dim.hidden

    formula_cells = set()
    for row in ws_src.iter_rows():
        for c in row:
            nc = ws.cell(row=c.row, column=c.column)
            nc.font = copy(c.font)                 # 跨工作簿须复制样式对象，不能用 _style 索引数组
            nc.fill = copy(c.fill)
            nc.border = copy(c.border)
            nc.alignment = copy(c.alignment)
            nc.protection = copy(c.protection)
            nc.number_format = c.number_format
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                v = ws_val.cell(row=c.row, column=c.column).value   # 旧缓存值占位（本期无数据的供应商保留）
                formula_cells.add((c.row, c.column))
            nc.value = v

    for rng in list(ws_src.merged_cells.ranges):
        ws.merge_cells(str(rng))
    if ws_src.freeze_panes:
        ws.freeze_panes = ws_src.freeze_panes
    wb_src.close()
    wb_val.close()
    return ws, formula_cells


def _summary_layout(ws):
    """解析「汇总」表头：返回 (不开票金额列, 金额列, 开票金额列, 总计列, 总计口径列)"""
    hdr = {}
    for j in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=j).value
        if v is not None:
            hdr[str(v).strip()] = j
    dcol = hdr.get("不开票金额")                      # 分公司 / 总部其他
    ccol = hdr.get("金额")                            # 简单型（单店 / 45 / 48 / 19 / 30）
    ecol = hdr.get("开票金额") or ccol
    tcol = hdr.get("总计")
    if dcol:
        # 分公司 E列=开票金额（总计ΣD）；总部其他 E列=金额（总计ΣE）
        basis = dcol if hdr.get("开票金额") else ecol
    else:
        basis = ccol
    return dcol, ccol, ecol, tcol, basis


def remove_absent_suppliers(ws, amounts):
    """删除「汇总」Sheet 中不在本期源文件夹供应商的整行（厂家名、收款信息、金额一并移除）。
    新建表格时调用：旧模板里含历史供应商，不在本期 amounts 中的整行删除，
    只保留源文件夹中存在的供应商。"""
    del_rows = []
    for i in range(2, ws.max_row + 1):
        name = ws.cell(row=i, column=1).value
        if name is None or not str(name).strip():
            continue
        if str(name).strip() in amounts:
            continue
        del_rows.append(i)
    for i in reversed(del_rows):   # 自下而上删除，避免行号偏移
        ws.delete_rows(i)


def _close(a, b):
    """单元格数值比较（容忍浮点尾差；None 视作 0）"""
    x = float(a) if is_number(a) else 0.0
    y = float(b) if is_number(b) else 0.0
    return abs(x - y) < 1e-6


def update_summary_amounts(ws, formula_cells, amounts, group):
    """按本期数据更新汇总Sheet金额；本表没有、本期有数据的供应商自动加行（收款信息留空）。
    amounts: {供应商: (不开票金额, 开票金额)}，均为不四舍五入值
    返回 (新增供应商列表, 金额有变化的供应商列表, 总计是否变化)"""
    dcol, ccol, ecol, tcol, basis = _summary_layout(ws)

    existing = {}
    for i in range(2, ws.max_row + 1):
        v = ws.cell(row=i, column=1).value
        if v is not None and str(v).strip():
            existing[str(v).strip()] = i

    added, changed = [], set()
    for sup, (no_inv, inv) in amounts.items():
        if sup in existing:
            i = existing[sup]
        else:
            i = ws.max_row + 1
            src_row = max(existing.values()) if existing else 1
            for j in range(1, ws.max_column + 1):
                sc, nc = ws.cell(row=src_row, column=j), ws.cell(row=i, column=j)
                nc.font = copy(sc.font)            # 同工作簿内复制样式对象
                nc.fill = copy(sc.fill)
                nc.border = copy(sc.border)
                nc.alignment = copy(sc.alignment)
                nc.protection = copy(sc.protection)
                nc.number_format = sc.number_format
            ws.cell(row=i, column=1, value=sup)
            existing[sup] = i
            added.append(sup)
            changed.add(sup)
        if dcol:
            if not _close(ws.cell(row=i, column=dcol).value, no_inv):
                changed.add(sup)
            ws.cell(row=i, column=dcol, value=no_inv)
            have_e = ws.cell(row=i, column=ecol).value is not None or (i, ecol) in formula_cells
            if have_e:
                if not _close(ws.cell(row=i, column=ecol).value, inv):
                    changed.add(sup)
                ws.cell(row=i, column=ecol, value=inv)
        elif ccol:
            target = inv if group == GROUP_SINGLE else no_inv
            if not _close(ws.cell(row=i, column=ccol).value, target):
                changed.add(sup)
            ws.cell(row=i, column=ccol, value=target)

    total_changed = False
    if tcol:
        total = sum((Decimal(str(ws.cell(row=i, column=basis).value))
                     for i in existing.values()
                     if is_number(ws.cell(row=i, column=basis).value)), Decimal("0"))
        cand = sorted(r for (r, c) in formula_cells if c == tcol)
        if not cand:                               # 已存在「汇总」（公式已转静态）：取该列最靠上的数值单元格
            cand = [i for i in range(2, ws.max_row + 1)
                    if is_number(ws.cell(row=i, column=tcol).value)]
        if cand:
            tcell = ws.cell(row=cand[0], column=tcol)
            total_changed = not _close(tcell.value, float(total))
            tcell.value = float(total)             # 写静态值（与旧文件一致）
    return added, [s for s in amounts if s in changed], total_changed


def _header_cols(ws):
    """解析「汇总」表头 → {列名: 列号}（行1）"""
    hdr = {}
    for j in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=j).value
        if v is not None:
            hdr[str(v).strip()] = j
    return hdr


def write_payment_info(ws, sup):
    """把配置里的「收款信息 / 不开票收款信息 / 备注(开票状态)」写入「汇总」Sheet 对应列。
    以界面配置为准覆盖旧值；未配置该供应商时跳过。返回是否有单元格被改动。"""
    cfg = supplier_config.get_supplier(sup)
    if cfg is None:
        return False
    hdr = _header_cols(ws)
    bank_col = hdr.get("收款信息")
    nobank_col = hdr.get("不开票收款信息")
    remark_col = hdr.get("备注")
    changed = False

    # 定位该供应商所在行
    row = None
    for i in range(2, ws.max_row + 1):
        v = ws.cell(row=i, column=1).value
        if v is not None and str(v).strip() == str(sup).strip():
            row = i
            break
    if row is None:
        return False

    def _set(col, val):
        nonlocal changed
        if col is None:
            return
        c = ws.cell(row=row, column=col)
        if c.value != val:
            c.value = val
            changed = True

    _set(bank_col, cfg.get("bank") or None)
    _set(nobank_col, cfg.get("no_bank") or None)
    status = cfg.get("invoice_status") or supplier_config.STATUS_NO_PLUS
    # 备注列写开票状态（三档文案）
    _set(remark_col, status)
    return changed


def _text_width(s):
    """文本显示宽度：中文按 2 字符计"""
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


def _disp_lines(text, width):
    """按列宽估算自动换行后的视觉行数"""
    eff = max(int(round(width)) - 1, 4)
    n = 0
    for line in text.split("\n"):
        n += max(1, -(-_text_width(line) // eff))
    return n


def fit_summary_columns(ws, max_width=55):
    """「汇总」列宽/行高自适应：列宽按最长一行内容（换行拆分、中文按 2 字符）计算并封顶，
    所有单元格启用自动换行；行高不足时按换行后行数抬高，确保内容完整显示。返回是否发生调整。"""
    changed = False

    # 1) 列宽（跳过整列为空的列）
    for j in range(1, ws.max_column + 1):
        need, has_value = 8, False
        for i in range(1, ws.max_row + 1):
            v = ws.cell(row=i, column=j).value
            if v is None:
                continue
            s = f"{v:.6f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)
            if s.startswith("="):
                continue
            has_value = True
            for line in s.split("\n"):
                need = max(need, _text_width(line) + 2)
        if not has_value:
            continue
        need = min(need, max_width)
        letter = get_column_letter(j)
        cur = ws.column_dimensions[letter].width
        if cur is None or abs(cur - need) > 0.4:
            ws.column_dimensions[letter].width = need
            changed = True

    # 2) 自动换行 + 行高不足时抬高
    for i in range(1, ws.max_row + 1):
        need_h = 0.0
        for j in range(1, ws.max_column + 1):
            c = ws.cell(row=i, column=j)
            v = c.value
            if v is None:
                continue
            s = f"{v:.6f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)
            if s.startswith("="):
                continue
            if not c.alignment.wrap_text:
                a = copy(c.alignment)
                a.wrap_text = True
                c.alignment = a
                changed = True
            w = ws.column_dimensions[get_column_letter(j)].width or 8.43
            n = _disp_lines(s, w)
            need_h = max(need_h, n * max(14.0, (c.font.size or 11) * 1.35))
        if need_h > 0:
            rd = ws.row_dimensions[i]
            if rd.height is not None and rd.height < need_h - 0.5:
                rd.height = round(need_h * 2) / 2
                changed = True
    return changed


# ============================================================
# 主流程
# ============================================================
def run_split(input_dir, output_dir):
    """核心拆分流程（供 CLI 与 GUI 复用）；日志走 print"""
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    print("=" * 60)
    print("  供应商对账单 - 按门店拆分 + 金额汇总")
    print("=" * 60)
    print(f"  源目录  : {input_dir}")
    print(f"  输出目录: {output_dir}")
    print()

    if not input_dir.exists():
        print(f"  [FAIL] 源目录不存在：{input_dir}")
        return 1

    files = sorted(p for p in input_dir.iterdir()
                   if p.suffix.lower() in (".xlsx", ".xlsm", ".xls") and not p.name.startswith("~$"))
    if not files:
        print("  [FAIL] 源目录下没有可处理的 xlsx 文件")
        return 1

    # 读取 + 归类
    print("1. 读取源文件并归类门店...")
    grouped = {}          # (输出文件, 供应商) → {rows, group, codes, months}
    supplier_order = []   # 供应商列顺序（按源文件首次出现）
    unclassified = set()
    total_rows = 0
    plat_rows = 0
    for f in files:
        recs = read_reconcile_rows(f)
        precs = read_platform_rows(f)
        print(f"   {f.name}: 有效对账行 {len(recs)}"
              + (f"，油菜花平台行 {len(precs)}" if precs else ""))
        for rec in recs + precs:
            fn, group, unc = route(rec["code"], rec["supplier"])
            if unc:
                unclassified.add(rec["code"])
            key = (fn, rec["supplier"])
            if key not in grouped:
                grouped[key] = {"rows": [], "group": group, "codes": set(), "months": set()}
                if rec["supplier"] not in supplier_order:
                    supplier_order.append(rec["supplier"])
            it = grouped[key]
            it["rows"].append(rec)
            it["codes"].add(rec["code"])
            if rec["month"]:
                it["months"].add(rec["month"])
            total_rows += 1
            if rec in precs:
                plat_rows += 1
    print(f"   合计有效行 {total_rows} 行（含平台 {plat_rows} 行），"
          f"{len(grouped)} 个（文件×供应商）组合")
    print()

    # 按输出文件分组，决定写入顺序：单独门店（按编号）→ 分公司 → 总部其他
    files_map = {}
    for (fn, sup), it in grouped.items():
        files_map.setdefault(fn, []).append((sup, it))

    def file_sort_key(fn):
        m = re.match(r"(\d+)号店", Path(fn).stem)
        if m:
            return (0, int(m.group(1)))
        if fn == BRANCH_FILE:
            return (1, 0)
        return (2, 0)

    file_order = sorted(files_map.keys(), key=file_sort_key)

    # 写入门店文件
    output_dir.mkdir(parents=True, exist_ok=True)
    print("2. 写入门店文件（含「汇总」Sheet 复制与金额更新）...")
    written, skipped, failed = [], [], []
    for fn in file_order:
        items = files_map[fn]
        out_path = output_dir / fn
        exists = out_path.exists()
        if exists:
            wb = load_workbook(out_path)
        else:
            wb = Workbook()
            wb.remove(wb.active)
        # 不开票供应商合法去向：总部其他、分公司、以及「单独拆分+不开票口径」的门店文件（如45/48）
        no_inv_ok = {HQ_OTHER_FILE, BRANCH_FILE}
        no_inv_ok.update(
            f"{c}号店.xlsx" for c, s in supplier_config.load_config().get("stores", {}).items()
            if s.get("group") == supplier_config.STORE_GROUP_SPLIT
            and s.get("invoice") == supplier_config.INVOICE_OFF)
        if fn not in no_inv_ok:
            stale = [s for s in wb.sheetnames if supplier_config.is_no_invoice(s)]
            if stale:
                print(f"   [WARN] {fn} 存在不开票供应商的旧 Sheet（{'、'.join(stale)}）："
                      f"按规则其数据不应出现在本文件，如需清理请手动删除该 Sheet")
        wrote_any = False
        for sup, it in items:
            sheet_name = clean_sheet_name(sup)
            if sheet_name in wb.sheetnames:
                print(f"   [SKIP] {fn} 已存在 Sheet「{sheet_name}」，跳过不覆盖")
                skipped.append(f"{fn} / {sheet_name}")
                continue
            write_supplier_sheet(wb, sheet_name, it["rows"], it["group"],
                                 period_str(it["months"]), supplier_rule(sup))
            wrote_any = True
            written.append(f"{fn} / {sheet_name}")

        # 「汇总」Sheet：不存在则从旧「按店铺对账」对应文件复制并更新金额；
        # 已存在则仅在金额与本期不一致时刷新（收款信息/备注等其余内容一律不动）
        summary_changed = False
        amounts = {sup: calc_unrounded(it["rows"], supplier_rule(sup))
                   for sup, it in items}
        payment_changed = False
        if "汇总" in wb.sheetnames:
            added, changed, tot_changed = update_summary_amounts(
                wb["汇总"], set(), amounts, items[0][1]["group"])
            if added or changed or tot_changed:
                summary_changed = True
                names = "、".join(dict.fromkeys(list(changed) + list(added)))
                print(f"   [OK]   {fn} 已刷新已有「汇总」金额（{names}）"
                      + ("，含新增供应商行（收款信息待补）" if added else ""))
            else:
                print(f"   [SKIP] {fn} 已有「汇总」，金额与本期一致")
            # 以界面配置为准：同步收款信息 / 开票状态备注
            for sup, _ in items:
                if write_payment_info(wb["汇总"], sup):
                    payment_changed = True
            if payment_changed:
                summary_changed = True
                print(f"   [OK]   {fn} 已按配置同步「汇总」收款信息/开票状态")
            if fit_summary_columns(wb["汇总"]):
                summary_changed = True
                print(f"   [OK]   {fn} 已自适应「汇总」列宽/行高（内容完整显示）")
        else:
            src = old_summary_source(fn)
            if src is None:
                print(f"   [WARN] {fn} 未找到旧汇总模板（{OLD_DIR}），本次不生成「汇总」Sheet")
            else:
                ws_sum, fcells = copy_summary_sheet(src, wb)
                remove_absent_suppliers(ws_sum, amounts)
                added, _, _ = update_summary_amounts(ws_sum, fcells, amounts, items[0][1]["group"])
                for sup, _ in items:
                    if write_payment_info(ws_sum, sup):
                        payment_changed = True
                fit_summary_columns(ws_sum)
                summary_changed = True
                print(f"   [OK]   {fn} 已复制旧「汇总」并按本期数据更新金额"
                      + (f"（新增供应商行：{'、'.join(added)}，收款信息待补）" if added else ""))

        if exists and not wrote_any and not summary_changed:
            print(f"   [SKIP] {fn} 本次无新 Sheet，文件未改动")
            continue
        try:
            wb.save(out_path)
            print(f"   [OK]   {fn} 已写入")
        except PermissionError:
            print(f"   [FAIL] {fn} 保存失败：文件被占用（已在 WPS/Excel 打开），请关闭后重跑")
            failed.append(fn)
    print()

    # 生成汇总
    print("3. 生成门店金额汇总...")
    summary_path = output_dir / SUMMARY_NAME
    try:
        write_summary(summary_path, file_order, files_map, supplier_order)
        print(f"   [OK]   {SUMMARY_NAME}")
    except PermissionError:
        alt = output_dir / f"{Path(SUMMARY_NAME).stem}_备份_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        write_summary(alt, file_order, files_map, supplier_order)
        print(f"   [WARN] {SUMMARY_NAME} 被占用，已写入备用文件：{alt.name}")
    print()

    # 统计
    print("=" * 60)
    print(f"完成：扫描 {len(files)} 个源文件、{total_rows} 行对账数据（含油菜花平台 {plat_rows} 行）")
    print(f"  输出文件 {len(file_order)} 个：新增 Sheet {len(written)} 个，跳过 {len(skipped)} 个，失败 {len(failed)} 个")
    if skipped:
        print("  跳过清单（已存在，未覆盖）：")
        for s in skipped:
            print(f"    - {s}")
    if failed:
        print("  失败清单（请关闭文件后重跑）：")
        for s in failed:
            print(f"    - {s}")
    if unclassified:
        print(f"  [WARN] 新门店未归类，已默认并入 {HQ_OTHER_FILE}："
              + "、".join(sorted(unclassified, key=int)) + "号店")
    print(f"  输出目录：{output_dir}")


def main():
    ap = argparse.ArgumentParser(description="按门店拆分供应商对账单并生成金额汇总")
    ap.add_argument("--input", default=INPUT_DIR, help="源对账文件目录")
    ap.add_argument("--output", default=OUTPUT_DIR, help="输出目录")
    args = ap.parse_args()
    run_split(args.input, args.output)


if __name__ == "__main__":
    try:
        main()
    except Exception:                      # exe 双击运行时不闪退，保留报错信息
        import traceback
        traceback.print_exc()
    finally:
        # 打包为 exe 后双击运行（无参数）时，结束后暂停，便于查看日志
        if getattr(sys, "frozen", False) and len(sys.argv) == 1:
            input("\n处理完成，按回车键退出...")