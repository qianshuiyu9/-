"""原创厂家专属处理管线 —— 独立于其他厂家."""
import os
import re
import pandas as pd
from config import OUTPUT_DIR, FIXED_INTEGRAL_MULTIPLIER, TARGET_COLUMNS

# 原创内部列名 → 我们的内部字段名
_YC_INTERNAL_COLS = [
    "序号", "类别", "_空1", "_空2", "_空3",
    "货名", "_空4", "_空5", "_空6", "_空7",
    "备注", "_空8", "_空9", "_空10", "_空11",
    "_空12", "仓库", "_空13", "货位", "_空14",
    "单位", "数量", "_空15", "_空16", "单价",
    "_空17", "金额", "_空18", "_空19", "条码",
]


def _parse_yc_file(filepath):
    """解析一个原创 xls/xlsx 文件，拆成多个销售单块.

    对 .xls 文件先 COM 转成 .xlsx（处理合并单元格），再用 openpyxl 读.
    返回 list of dict:
        {
            "客户名": str,
            "应收": float/str,
            "实收": float/str,
            "货名列": list[str],
            "类别列": list[str],
            "数量列": list[float],
            "单价列": list[float],
            "单位列": list[str],
            "条码列": list[str],
            "备注列": list[str],
        }
    """
    import tempfile
    from openpyxl import load_workbook

    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".xls":
        tmp_xlsx = os.path.join(tempfile.gettempdir(), "_yc_conv_" + str(os.getpid()) + ".xlsx")
        from excel_processor import _convert_xls_to_xlsx
        converted = _convert_xls_to_xlsx(os.path.abspath(filepath), tmp_xlsx)
        if not converted:
            print(f"  [警告] 无法转换 {filepath}，跳过")
            return []
        wb = load_workbook(tmp_xlsx, data_only=True)
        try:
            os.remove(tmp_xlsx)
        except OSError:
            pass
    else:
        wb = load_workbook(filepath, data_only=True)

    ws = wb.active
    rows = ws.max_row
    blocks = []
    current = None

    for r in range(1, rows + 1):
        def cv(col):
            """1-based 取单元格值，None → ''."""
            v = ws.cell(r, col).value
            return v

        # 客户行：C1 = "客户："，客户名在 C5
        v1 = cv(1)
        v5 = cv(5)
        cell0 = str(v1).strip() if v1 is not None else ""
        cell5 = str(v5).strip() if v5 is not None else ""

        if cell0 == "客户：" and cell5:
            if current and current.get("货名列"):
                blocks.append(current)
            current = {
                "客户名": cell5,
                "应收": "",
                "实收": "",
                "货名列": [],
                "类别列": [],
                "数量列": [],
                "单价列": [],
                "单位列": [],
                "条码列": [],
                "备注列": [],
            }
            continue

        if current is None:
            continue

        # 金额行：C8="应收：" 且 C23="实收："（注意 C1 是空字符串）
        v8 = cv(8)
        v23 = cv(23)
        cell8 = str(v8).strip() if v8 is not None else ""
        cell23 = str(v23).strip() if v23 is not None else ""

        if cell8 == "应收：" and cell23 == "实收：":
            ar = cv(10)    # 应收值在 C10
            paid = cv(26)  # 实收值在 C26
            current["应收"] = float(ar) if ar is not None and str(ar).strip() else (str(ar).strip() if ar is not None else "")
            current["实收"] = float(paid) if paid is not None and str(paid).strip() else (str(paid).strip() if paid is not None else "")
            continue

        # 跳过合计行、付款方式行、表头行
        if cell0 in ("合计", "付款方式", "序"):
            continue

        # 商品行：C1 是数字序号
        if not cell0:
            continue
        try:
            seq = float(cell0)
            if pd.isna(seq):
                continue
        except (ValueError, TypeError):
            continue
        if not (seq == int(seq)):
            continue

        # 商品数据列（openpyxl 1-based，COM 转换后）
        name = str(cv(6)).strip() if cv(6) is not None else ""     # 货名 C6
        category = str(cv(2)).strip() if cv(2) is not None else ""   # 类别 C2
        qty_v = cv(22)    # 数量 C22
        price_v = cv(25)  # 单价 C25
        unit = str(cv(21)).strip() if cv(21) is not None else ""     # 单位 C21
        barcode = str(cv(30)).strip() if cv(30) is not None else ""  # 条码 C30
        remark = str(cv(11)).strip() if cv(11) is not None else ""   # 备注 C11

        if not name or name == "nan":
            continue

        try:
            qty = float(qty_v) if qty_v is not None and str(qty_v).strip() not in ("", "nan") else 0
        except (ValueError, TypeError):
            qty = 0
        try:
            price = float(price_v) if price_v is not None and str(price_v).strip() not in ("", "nan") else 0
        except (ValueError, TypeError):
            price = 0

        current["货名列"].append(name)
        current["类别列"].append(category)
        current["数量列"].append(qty)
        current["单价列"].append(price)
        current["单位列"].append(unit)
        current["条码列"].append(barcode)
        current["备注列"].append(remark)

    if current and current.get("货名列"):
        blocks.append(current)

    wb.close()
    return blocks


def _yc_clean_display_name(name):
    """原创专属展示名清理.

    规则：
    1. 去掉头部括号内容：（整盒出12个装）XX → XX
    2. 去掉尾部括号规格：XX（288/箱）→ XX
    3. 去掉尾部斜杠规格：XX展示盒/12/只 → XX展示盒
    4. 保留商品编码中的横线和数字
    """
    # 0) 反复去掉括号组前面的裸中文前缀词（如"特供款"、"活动款"等标记词）
    _YC_KEEP_PREFIXES = {"盲盒", "玩具", "包包", "杯", "碗", "盒", "碟", "盘"}  # 疑似商品名的词不删
    while True:
        prev = name
        m = re.match(r"^\s*([\u4e00-\u9fff]+)\s*[（(]", name)
        if m:
            prefix = m.group(1)
            if prefix not in _YC_KEEP_PREFIXES:
                name = "（" + name[m.end(1):]
                # 即：把 "特供款（..." 变成 "（..."
        if name == prev:
            break

    # 1) 反复去掉开头的括号组（整盒出N个装 / 整盒出N个装）等
    changed = True
    while changed:
        changed = False
        # 半角或全角括号的"整盒出N个装"类
        m = re.match(r"^\s*[（(].*?[）)]\s*", name)
        if m:
            name = name[m.end():]
            changed = True

    # 2) 去掉尾部括号（规格/装箱信息）
    for _ in range(3):
        prev = name
        name = re.sub(r"\s*[（(].*?[）)]\s*$", "", name)
        if name == prev:
            break

    # 3) 去掉尾部斜杠规格：XX展示盒/12/只 或 XX/12只 或 XX/12/
    #    先处理连续 /N/单位 的情况
    for _ in range(3):
        prev = name
        # /N/单位 如 /12/只 /12/件
        name = re.sub(r"\s*/\s*\d+\s*/\s*[^\s/]+?\s*$", "", name)
        # 尾部 /N单位 如 /12只 /12个
        name = re.sub(r"\s*/\s*\d+\s*[^\s/]+?\s*$", "", name)
        # 尾部单独 /单位 如 /只 /件
        name = re.sub(r"\s*/\s*[^\s/]+?\s*$", "", name)
        if name == prev:
            break

    # 4) 清理残留空白、标点
    name = re.sub(r"\s+", " ", name).strip()
    name = re.sub(r"[，,。.、；;：:]+$", "", name)
    name = re.sub(r"^[，,。.、；;：:]+", "", name)
    return name.strip()


def _yc_classify_spec_tag(category):
    """原创：根据类别列的值判断规格和标签.

    规则：
    - 类别 == "玩具" → 玩具类 / 儿童玩具
    - 类别是其他 → 生活类，按家家乐的分类逻辑：
        水杯类：陶瓷杯 / 玻璃杯 / 马克杯 / 保温杯 / 杯
        餐具类：碗 / 陶瓷碗 / 餐具 / 盘 / 碟 / 筷 / 勺
        数码：数码 / 充电 / 耳机 / 音箱
        箱包：包包 / 时尚小包 / 手提包 / 包 / 袋 / 箱
        挂件：挂件 / 挂饰
        其他：生活用品
    """
    cat = str(category).strip()

    if cat == "玩具":
        return "玩具类", "儿童玩具"

    # 生活类 — 先处理可能含多个关键词的类别
    cup_kws = ["陶瓷杯", "玻璃杯", "马克杯", "保温杯", "水杯", "杯"]
    tableware_kws = ["陶瓷碗", "餐具", "泡面碗", "碗", "盘", "碟", "筷", "勺"]
    digital_kws = ["数码", "充电", "耳机", "音箱", "数据线"]
    bag_kws = ["时尚小包", "手提包", "包包", "包", "袋", "箱"]
    pendant_kws = ["挂件", "挂饰"]

    # 先长后短匹配，避免"陶瓷杯"被"杯"抢先
    for kw in cup_kws:
        if kw in cat:
            return "生活类", "水杯"
    for kw in tableware_kws:
        if kw in cat:
            return "生活类", "餐具"
    for kw in digital_kws:
        if kw in cat:
            return "生活类", "数码"
    for kw in pendant_kws:
        if kw in cat:
            return "生活类", "挂件"
    for kw in bag_kws:
        if kw in cat:
            return "生活类", "箱包"

    return "生活类", "生活用品"


def _yc_extract_box_size(name):
    """从原创货品名中提取整盒的每盒个数 N.

    匹配模式：整盒出N个装 / 整盒N个装 / 整盒出N装 / 整盒N装 等
    返回 N（int）或 None
    """
    # 整盒出12个装 / 整盒出6个装 / 整盒12个装
    m = re.search(r"整盒.*?(\d+)\s*个?\s*装", name)
    if m:
        return int(m.group(1))
    # 整盒出N （少数情况）
    m = re.search(r"整盒.*?(\d+)", name)
    if m:
        return int(m.group(1))
    return None


def process_yuanchuang_all():
    """处理所有原创文件，按客户名合并输出."""
    input_dir = os.path.join(
        os.path.join(os.path.expanduser("~"), "Desktop"),
        "原始数据文件", "原创"
    )
    if not os.path.isdir(input_dir):
        print(f"[原创] 输入目录不存在: {input_dir}")
        return []

    # 清空原创输出目录
    out_dir = os.path.join(OUTPUT_DIR, "原创")
    os.makedirs(out_dir, exist_ok=True)
    import shutil
    for f in os.listdir(out_dir):
        if f.startswith("~$"):
            continue  # Excel 临时锁定文件，跳过
        fp = os.path.join(out_dir, f)
        if os.path.isfile(fp):
            try:
                os.remove(fp)
            except OSError:
                pass

    # 收集所有块
    import glob as _glob
    all_files = []
    for ext in ["*.xls", "*.xlsx"]:
        all_files.extend(_glob.glob(os.path.join(input_dir, ext)))
    print(f"\n[原创] 发现 {len(all_files)} 个文件")

    # 按客户名聚合数据：客户 → list of rows
    customer_rows = {}    # customer_name → list of dict
    customer_ar_sum = {}  # customer_name → 应收汇总（float）
    customer_paid_sum = {}  # customer_name → 实收汇总（float）

    for fp in sorted(all_files):
        blocks = _parse_yc_file(fp)
        print(f"  {os.path.basename(fp)} → {len(blocks)} 个销售单块")
        for block in blocks:
            cust = block["客户名"]
            if cust not in customer_rows:
                customer_rows[cust] = []
                customer_ar_sum[cust] = 0.0
                customer_paid_sum[cust] = 0.0

            # 累加应收/实收（只有是数字才加）
            ar = block["应收"]
            paid = block["实收"]
            try:
                customer_ar_sum[cust] += float(ar)
            except (ValueError, TypeError):
                pass
            try:
                customer_paid_sum[cust] += float(paid)
            except (ValueError, TypeError):
                pass

            for i, name in enumerate(block["货名列"]):
                qty = block["数量列"][i]
                price = block["单价列"][i]

                # === 拆解逻辑 ===
                box_n = _yc_extract_box_size(name)
                decomposed = False
                if box_n is not None and "整盒" in name and qty > 0:
                    if qty % box_n != 0:
                        decomposed = True
                        new_qty = qty * box_n
                        new_price = round(price / box_n, 2) if price > 0 else price
                    else:
                        new_qty = qty
                        new_price = price
                else:
                    new_qty = qty
                    new_price = price

                display_name = _yc_clean_display_name(name)
                spec_val, tag_val = _yc_classify_spec_tag(block["类别列"][i])

                customer_rows[cust].append({
                    "原始名称": name,
                    "名称": name,
                    "类别": block["类别列"][i],
                    "单位": "个",
                    "进货数量": new_qty,
                    "单价": new_price,
                    "条码": block["条码列"][i],
                    "规格": spec_val,
                    "用途": "兑换, 零售",
                    "标签": tag_val,
                    "积分倍数": FIXED_INTEGRAL_MULTIPLIER,
                    "展示名": display_name if display_name else name,
                    "_decomposed": decomposed,
                })

    print(f"\n[原创] 合并后共 {len(customer_rows)} 个客户")

    # 输出一个客户一个文件
    results = []
    for cust, rows in customer_rows.items():
        # 去重：原始名称 + 单价相同的合并（数量求和）
        df = pd.DataFrame(rows)
        if len(df) > 1:
            df = _yc_deduplicate(df)

        # 过滤空数量/空名称
        df = df[df["进货数量"].notna() & (df["进货数量"] != 0) & df["名称"].notna() & (df["名称"] != "")]

        # 列顺序：TARGET_COLUMNS + 应收金额 + 实收金额
        final_cols = list(TARGET_COLUMNS) + ["应收金额", "实收金额"]
        for c in final_cols:
            if c not in df.columns:
                df[c] = ""
        df = df[final_cols]

        # 应收/实收汇总：只在第一行填值，其他行空着
        ar_total = round(customer_ar_sum.get(cust, 0), 2) if customer_ar_sum.get(cust, 0) else ""
        paid_total = round(customer_paid_sum.get(cust, 0), 2) if customer_paid_sum.get(cust, 0) else ""

        # 文件名：原创+客户名（清理非法字符）
        safe_cust = re.sub(r"[\\/:*?\"<>|\n\r\t]+", "_", cust).strip()
        out_name = f"原创{safe_cust}.xlsx"
        out_path = os.path.join(out_dir, out_name)

        # 用 openpyxl 写入（原创无图片，简化）
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        # 表头
        for c_idx, col_name in enumerate(final_cols, start=1):
            ws.cell(row=1, column=c_idx, value=col_name)

        # 数据行
        ar_col_idx = final_cols.index("应收金额") + 1
        paid_col_idx = final_cols.index("实收金额") + 1
        for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
            for c_idx, col_name in enumerate(final_cols, start=1):
                val = row.get(col_name, "")
                # 应收/实收：只在第一行填汇总，其他行空
                if col_name == "应收金额":
                    val = ar_total if r_idx == 2 else ""
                elif col_name == "实收金额":
                    val = paid_total if r_idx == 2 else ""
                # 清理 NaN
                if isinstance(val, float) and pd.isna(val):
                    val = ""
                ws.cell(row=r_idx, column=c_idx, value=val)

        wb.save(out_path)

        decomp_count = sum(1 for r in rows if r["_decomposed"])
        print(f"  ✓ {out_name}: {len(df)} 行（拆解 {decomp_count} 行）")
        results.append({"name": out_name, "customer": cust, "rows": len(df), "decomposed": decomp_count})

    return results


def _yc_deduplicate(df):
    """原创去重：原始名称 + 单价相同的合并，数量求和."""
    if len(df) <= 1:
        return df
    df = df.copy()
    # 把原始名称和单价拼成 key
    df["_key"] = df["原始名称"].astype(str) + "||" + df["单价"].astype(str)

    grouped = df.groupby("_key", sort=False)
    merged = []
    for key, g in grouped:
        row = g.iloc[0].copy()
        row["进货数量"] = g["进货数量"].sum()
        merged.append(row)

    result = pd.DataFrame(merged)
    result = result.drop(columns=["_key", "_decomposed", "原始名称"], errors="ignore")
    return result.reset_index(drop=True)


if __name__ == "__main__":
    process_yuanchuang_all()
