"""原创厂家专属处理管线 —— 适配新单店单表格式（含图片）."""
import os
import re
import tempfile
import pandas as pd
from config import OUTPUT_DIR, FIXED_INTEGRAL_MULTIPLIER, TARGET_COLUMNS


def _parse_yc_file(filepath):
    """解析一个原创 xls/xlsx 文件."""
    from openpyxl import load_workbook
    from excel_processor import _convert_xls_to_xlsx, _read_images_from_wb

    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".xls":
        tmp_xlsx = os.path.join(tempfile.gettempdir(), "_yc_conv_" + str(os.getpid()) + ".xlsx")
        converted = _convert_xls_to_xlsx(os.path.abspath(filepath), tmp_xlsx)
        if not converted:
            print(f"  [警告] 无法转换 {filepath}，跳过")
            return None
        xlsx_path = tmp_xlsx
    else:
        xlsx_path = filepath

    try:
        df_raw = pd.read_excel(xlsx_path, header=None, dtype=object)
        customer_name = ""
        if len(df_raw) > 1:
            v = df_raw.iloc[1, 4]
            if v is not None and str(v).strip():
                customer_name = str(v).strip()

        df = pd.read_excel(xlsx_path, header=3, dtype=object)
        if "行号" in df.columns:
            df = df[df["行号"].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x))]
        df = df.reset_index(drop=True)

        wb = load_workbook(xlsx_path)
        wb_img_map = _read_images_from_wb(wb)
        wb.close()

        img_map = {}
        for opxl_row, img_obj in wb_img_map.items():
            pd_idx = opxl_row - 5
            if 0 <= pd_idx < len(df):
                img_map[opxl_row] = img_obj

        return {"客户名": customer_name, "df": df, "img_map": img_map}

    finally:
        if ext == ".xls" and os.path.exists(xlsx_path):
            try:
                os.remove(xlsx_path)
            except OSError:
                pass


def _yc_clean_display_name(name):
    """原创展示名清理：去掉尾部的 (N/箱)、规格括号等."""
    # 去掉尾部 (N/箱)、（N/箱）、(N/盒) 等
    for _ in range(3):
        prev = name
        name = re.sub(r"\s*[（(]\s*\d+\s*/\s*箱\s*[）)]\s*$", "", name)
        name = re.sub(r"\s*[（(]\s*\d+\s*/\s*盒\s*[）)]\s*$", "", name)
        if name == prev:
            break

    # 去掉尾部其他括号组（含完整/半角括号）
    for _ in range(3):
        prev = name
        name = re.sub(r"\s*[（(].*?[）)]\s*$", "", name)
        if name == prev:
            break

    name = re.sub(r"\s+", " ", name).strip()
    name = re.sub(r"[，,。.、；;：:]+$", "", name)
    return name.strip()


def _yc_classify_spec_tag(category):
    """根据商品分类值判断规格和标签.

    匹配顺序（三层优先级）：
      1. toy_exact — 玩具精确词（棋盘/扑克牌等，两字以上，不会和硬分类冲突）
      2. 硬分类    — 水杯 → 餐具(含盘) → 数码 → 挂件 → 箱包
      3. toy_fallback — 玩具泛词兜底
      4. 默认生活类/生活用品
    """
    cat = str(category).strip()

    # --- 第 1 层：玩具精确词（两字以上完整词，不会和硬分类冲突）---
    toy_exact = ["棋盘", "象棋", "飞行棋", "斗兽棋", "军棋", "围棋", "跳棋",
                 "五子棋", "大富翁", "扑克牌", "卡牌", "纸牌"]
    for kw in toy_exact:
        if kw in cat:
            return "玩具类", "儿童玩具"

    # --- 第 2 层：硬分类 ---
    cup_kws = ["陶瓷杯", "玻璃杯", "马克杯", "保温杯", "水杯", "杯"]
    tableware_kws = ["陶瓷碗", "餐具", "泡面碗", "碗盘", "碗", "盘", "碟", "筷", "勺"]
    digital_kws = ["数码", "充电", "耳机", "音箱", "数据线"]
    bag_kws = ["时尚小包", "手提包", "包包", "皮包", "包", "袋", "箱"]
    pendant_kws = ["挂件", "挂饰"]
    decor_kws = ["水晶球"]

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
    for kw in decor_kws:
        if kw in cat:
            return "生活类", "摆件"
    for kw in bag_kws:
        if kw in cat:
            return "生活类", "箱包"

    # --- 第 3 层：玩具泛词兜底 ---
    toy_fallback = ["玩具", "盲盒", "公仔", "娃娃", "积木", "棋", "游戏"]
    for kw in toy_fallback:
        if kw in cat:
            return "玩具类", "儿童玩具"

    return "生活类", "生活用品"


def _yc_extract_pack_size(name):
    """从货品名提取包装内个数 N.

    匹配格式：
      旧格式：整盒出36个装 / 整盒36
      新格式：(36/箱) / （60/箱） / 40/箱
    返回 N 或 None.
    """
    # 新格式（优先，更精确）
    m = re.search(r"[（(]\s*(\d+)\s*/\s*[箱盒包袋]\s*[）)]", name)
    if m:
        return int(m.group(1))
    # 无括号的新格式
    m = re.search(r"(?<!\d)(\d+)\s*/\s*[箱盒包袋](?!\d)", name)
    if m:
        return int(m.group(1))
    # 旧格式
    m = re.search(r"整盒.*?(\d+)\s*个?\s*装", name)
    if m:
        return int(m.group(1))
    m = re.search(r"整盒.*?(\d+)", name)
    if m:
        return int(m.group(1))
    return None


def _round_half_up(v, ndigits=2):
    from decimal import Decimal, ROUND_HALF_UP
    d = Decimal(str(v)).quantize(Decimal("0." + "0" * ndigits), rounding=ROUND_HALF_UP)
    return float(d)


def process_yuanchuang_all(suppliers=None):
    """处理所有原创文件，每个文件一个客户，输出一个文件.

    suppliers: list[str] 或 None。若给定且不包含 '原创'，直接返回 [].
    """
    if suppliers is not None and "原创" not in suppliers:
        return []

    input_dir = os.path.join(
        os.path.join(os.path.expanduser("~"), "Desktop"),
        "原始数据文件", "原创"
    )
    if not os.path.isdir(input_dir):
        print(f"[原创] 输入目录不存在: {input_dir}")
        return []

    out_dir = os.path.join(OUTPUT_DIR, "原创")
    os.makedirs(out_dir, exist_ok=True)
    import shutil
    for f in os.listdir(out_dir):
        fp = os.path.join(out_dir, f)
        if os.path.isfile(fp) and not f.startswith("~$"):
            try:
                os.remove(fp)
            except OSError:
                pass

    import glob as _glob
    all_files = []
    for ext in ["*.xls", "*.xlsx"]:
        all_files.extend(_glob.glob(os.path.join(input_dir, ext)))
    print(f"\n[原创] 发现 {len(all_files)} 个文件")

    results = []

    for fp in sorted(all_files):
        parsed = _parse_yc_file(fp)
        if parsed is None:
            continue

        cust = parsed["客户名"] or os.path.splitext(os.path.basename(fp))[0]
        df = parsed["df"]
        img_map = parsed["img_map"]

        print(f"  {os.path.basename(fp)} → 客户={cust!r}, 商品行={len(df)}, 图片={len(img_map)}")

        if len(df) == 0:
            print(f"    [跳过] 无商品数据")
            continue

        rows_data = []
        for pd_idx, row in df.iterrows():
            name = str(row.get("商品名称", "")).strip()
            if not name or name == "nan":
                continue

            category = str(row.get("商品分类", "")).strip()

            # 数量、单价、金额
            qty_v = row.get("数量")
            price_v = row.get("单价")
            amount_v = row.get("金额")
            try:
                qty = float(qty_v) if pd.notna(qty_v) and str(qty_v).strip() not in ("", "nan") else 0
            except (ValueError, TypeError):
                qty = 0
            try:
                price = float(price_v) if pd.notna(price_v) and str(price_v).strip() not in ("", "nan") else 0
            except (ValueError, TypeError):
                price = 0

            barcode_v = row.get("商品条码")
            barcode = str(barcode_v).strip() if pd.notna(barcode_v) else ""

            sku_v = row.get("商品编号")
            sku = str(sku_v).strip() if pd.notna(sku_v) else ""

            # === 规格拆解（沿用原创旧规则）===
            # 只匹配"整盒出N个装"格式，且数量不是 N 的倍数时才拆解
            box_n = None
            m = re.search(r"整盒出?\s*(\d+)\s*个?\s*装", name)
            if m:
                box_n = int(m.group(1))
            if box_n is None:
                m = re.search(r"整盒(\d+)", name)
                if m:
                    box_n = int(m.group(1))

            decomposed = False
            if box_n is not None and box_n > 1 and qty > 0 and price > 0:
                if qty % box_n != 0:
                    decomposed = True
                    new_qty = qty * box_n
                    new_price = _round_half_up(price / box_n, 2)
                else:
                    new_qty = qty
                    new_price = price
            else:
                new_qty = qty
                new_price = price

            display_name = _yc_clean_display_name(name)
            spec_val, tag_val = _yc_classify_spec_tag(category)

            rows_data.append(({
                "类别": category,               # 商品分类原始值（皮包系列、碗盘组合类等）
                "原始名称": name,
                "名称": name,
                "单位": "个",
                "进货数量": new_qty,
                "单价": new_price,
                "条码": barcode,
                "规格": spec_val,
                "用途": "兑换, 零售",
                "标签": tag_val,
                "积分倍数": FIXED_INTEGRAL_MULTIPLIER,
                "供应商商品编码": "",            # 原创不填供应商编码
                "尺寸": "",
                "展示名": display_name if display_name else name,
                "总金额": "",                    # 金额列暂空（旧表应收/实收逻辑已废弃）
                "_decomposed": decomposed,
            }, pd_idx))

        if not rows_data:
            print(f"    [跳过] 无有效商品行")
            continue

        all_rows = [r[0] for r in rows_data]
        all_indices = [r[1] for r in rows_data]

        out_df = pd.DataFrame(all_rows)
        orig_indices = list(range(len(out_df)))

        mask = out_df["进货数量"].notna() & (out_df["进货数量"] != 0) & out_df["名称"].notna() & (out_df["名称"] != "")
        out_df = out_df[mask].reset_index(drop=True)
        orig_indices = [orig_indices[i] for i in range(len(mask)) if mask.iloc[i]]

        final_cols = list(TARGET_COLUMNS)
        for c in final_cols:
            if c not in out_df.columns:
                out_df[c] = ""
        out_df = out_df[final_cols]

        safe_cust = re.sub(r"[\\/:*?\"<>|\n\r\t]+", "_", cust).strip()
        out_name = f"原创{safe_cust}.xlsx"
        out_path = os.path.join(out_dir, out_name)

        from excel_processor import _write_output_with_images

        row_to_header_row = {}
        for out_i, pd_i in enumerate(orig_indices):
            row_to_header_row[out_i] = pd_i + 5

        _write_output_with_images(out_df, img_map, row_to_header_row, out_path)

        decomp_count = sum(1 for rd in rows_data if rd[0]["_decomposed"])
        img_count = len([v for v in row_to_header_row.values() if v in img_map])
        print(f"  ✓ {out_name}: {len(out_df)} 行（拆解 {decomp_count} 行, 图片 {img_count}）")
        results.append({"name": out_name, "customer": cust, "rows": len(out_df), "spec_changes": decomp_count, "image_count": img_count})

    return results


if __name__ == "__main__":
    process_yuanchuang_all()
