"""原创厂家专属处理管线 —— 适配新单店单表格式（含图片）."""
import os
import re
import tempfile
import pandas as pd
from config import OUTPUT_DIR, FIXED_INTEGRAL_MULTIPLIER, TARGET_COLUMNS


def _cell(row, c):
    """安全读取一行中的某列，空 / NaN 返回空字符串."""
    try:
        v = row[c]
    except (IndexError, KeyError, TypeError):
        return ""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _cell_num(row, c):
    """读取数值单元格，无法转换返回 None."""
    s = _cell(row, c)
    try:
        return float(s)
    except ValueError:
        return None


def _is_row_index(v):
    """商品行首列是否为序号（数字）."""
    s = str(v).strip() if v is not None else ""
    if s in ("", "nan", "None"):
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def _parse_yc_old(df_raw):
    """解析老版「原创生活馆 销售单」格式（一个文件可含多个客户块）.

    版式：客户：→ 序(表头) → 商品行 → 合计/应收实收 → 客户签名
    列：1=类别 5=货名 20=单位 21=数量 24=单价 26=金额 29=条码
    """
    nrows, ncols = df_raw.shape
    blocks = []
    i = 0
    while i < nrows:
        if _cell(df_raw.iloc[i], 0) != "客户：":
            i += 1
            continue

        # 客户名：同行第 1~8 列第一个非空值（跳过"客户电话："等标签）
        cust = ""
        for c in range(1, min(ncols, 9)):
            v = _cell(df_raw.iloc[i], c)
            if v and not v.endswith("："):
                cust = v
                break

        # 表头行（第 0 列为"序"）
        hdr = None
        for j in range(i + 1, min(i + 8, nrows)):
            if _cell(df_raw.iloc[j], 0) == "序":
                hdr = j
                break
        if hdr is None:
            i += 1
            continue

        recs = []
        total = None
        j = hdr + 1
        while j < nrows:
            first = _cell(df_raw.iloc[j], 0)
            if first == "客户：" or first.startswith("客户签名"):
                break
            # 应收行："应收："右侧第一个数值
            if total is None:
                recv_col = next(
                    (c for c in range(ncols) if _cell(df_raw.iloc[j], c) == "应收："), None
                )
                if recv_col is not None:
                    for c in range(recv_col + 1, ncols):
                        v = _cell_num(df_raw.iloc[j], c)
                        if v is not None:
                            total = v
                            break
            if _is_row_index(first):
                name = _cell(df_raw.iloc[j], 5)
                if name:
                    recs.append({
                        "商品分类": _cell(df_raw.iloc[j], 1),
                        "商品名称": name,
                        "单位": _cell(df_raw.iloc[j], 20),
                        "数量": _cell_num(df_raw.iloc[j], 21),
                        "单价": _cell_num(df_raw.iloc[j], 24),
                        "金额": _cell_num(df_raw.iloc[j], 26),
                        "商品条码": _cell(df_raw.iloc[j], 29),
                        "商品编号": "",
                    })
            j += 1

        if recs:
            blocks.append({"客户名": cust, "df": pd.DataFrame(recs), "img_map": {}, "合计金额": total})
        i = j
    return blocks


def _parse_yc_file(filepath):
    """解析一个原创 xls/xlsx 文件，返回订单列表.

    新版「商品销售报价」：每文件 1 单。
    老版「原创生活馆 销售单」：每文件可能多块，同客户合并为 1 单。
    """
    from openpyxl import load_workbook
    from excel_processor import _convert_xls_to_xlsx, _read_images_from_wb

    ext = os.path.splitext(filepath)[1].lower()
    tmp_xlsx = None

    if ext == ".xls":
        tmp_xlsx = os.path.join(tempfile.gettempdir(), "_yc_conv_" + str(os.getpid()) + ".xlsx")
        converted = _convert_xls_to_xlsx(os.path.abspath(filepath), tmp_xlsx)
        if not converted:
            print(f"  [警告] 无法转换 {filepath}，跳过")
            return []
        xlsx_path = tmp_xlsx
    else:
        xlsx_path = filepath

    try:
        df_raw = pd.read_excel(xlsx_path, header=None, dtype=object)
        head_n = min(15, len(df_raw))
        is_new = any(_cell(df_raw.iloc[i], 0) == "行号" for i in range(head_n))
        is_old = (not is_new) and any(
            _cell(df_raw.iloc[i], 0) == "序" and _cell(df_raw.iloc[i], 5) == "货名"
            for i in range(head_n)
        )

        if is_old:
            blocks = _parse_yc_old(df_raw)
            if not blocks:
                print(f"  [警告] 老格式未解析出商品行: {os.path.basename(filepath)}")
                return []
            # 同客户多块合并为一个订单
            merged = {}
            order = []
            for b in blocks:
                key = b["客户名"]
                if key not in merged:
                    merged[key] = b
                    order.append(key)
                else:
                    m = merged[key]
                    m["df"] = pd.concat([m["df"], b["df"]], ignore_index=True)
                    if b["合计金额"] is not None:
                        m["合计金额"] = (m["合计金额"] or 0) + b["合计金额"]
            print(f"  [老格式] {os.path.basename(filepath)}：{len(blocks)} 个销售块 → {len(order)} 个客户")
            return [merged[k] for k in order]

        if not is_new:
            print(f"  [警告] 无法识别版式，跳过: {os.path.basename(filepath)}")
            return []

        customer_name = ""
        if len(df_raw) > 1:
            v = df_raw.iloc[1, 4]
            if v is not None and str(v).strip():
                customer_name = str(v).strip()

        # 合计金额："合计"行在"金额"列上的值（用于输出文件名）
        total_amount = None
        hdr_idx = None
        for i in range(min(10, len(df_raw))):
            if str(df_raw.iloc[i, 0]).strip() == "行号":
                hdr_idx = i
                break
        if hdr_idx is not None:
            amount_col = None
            for c in range(df_raw.shape[1]):
                if str(df_raw.iloc[hdr_idx, c]).strip() == "金额":
                    amount_col = c
                    break
            if amount_col is not None:
                for i in range(hdr_idx + 1, len(df_raw)):
                    if str(df_raw.iloc[i, 0]).strip() == "合计":
                        v = df_raw.iloc[i, amount_col]
                        if pd.notna(v):
                            total_amount = v
                        break

        df = pd.read_excel(xlsx_path, header=3, dtype=object)
        if "行号" in df.columns:
            df = df[df["行号"].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x))]
        df = df.reset_index(drop=True)

        wb = load_workbook(xlsx_path)
        wb_img_map = _read_images_from_wb(wb)

        # wb close 之前把 openpyxl Image 对象转成纯 bytes，否则 wb.close() 后 _data() 失效
        img_map_raw = {}
        for opxl_row, img_obj in wb_img_map.items():
            try:
                img_map_raw[opxl_row] = img_obj._data() if hasattr(img_obj, '_data') else img_obj
            except Exception:
                pass
        wb.close()

        img_map = {}
        for opxl_row, raw_bytes in img_map_raw.items():
            pd_idx = opxl_row - 5
            if 0 <= pd_idx < len(df):
                img_map[opxl_row] = raw_bytes

        return [{
            "客户名": customer_name,
            "df": df,
            "img_map": img_map,
            "合计金额": total_amount,
        }]

    finally:
        if ext == ".xls" and tmp_xlsx and os.path.exists(tmp_xlsx):
            try:
                os.remove(tmp_xlsx)
            except OSError:
                pass


def _yc_clean_display_name(name):
    """原创展示名清理：去掉整盒说明、(N/箱)规格括号等."""
    # 0) 去掉任意位置的"整盒"说明括号（最优先）
    name = re.sub(r"[（(]\s*整盒[^（）()]*[）)]", "", name)

    # 1) 去掉任意位置的 (N/箱)、(N/盒)、(N入) 规格括号（不再只限尾部）
    for _ in range(3):
        prev = name
        name = re.sub(r"\s*[（(]\s*\d+\s*/\s*箱\s*[）)]", "", name)
        name = re.sub(r"\s*[（(]\s*\d+\s*/\s*盒\s*[）)]", "", name)
        name = re.sub(r"\s*[（(]\s*\d+\s*入\s*[）)]", "", name)
        if name == prev:
            break

    # 2) 去掉尾部其他括号组（品牌/系列括号：如"（蘑菇兔）"保留内容，空括号去掉）
    for _ in range(3):
        prev = name
        name = re.sub(r"\s*[（(].*?[）)]\s*$", "", name)
        if name == prev:
            break

    name = re.sub(r"\s+", " ", name).strip()
    name = re.sub(r"[，,。.、；;：:]+$", "", name)
    return name.strip()


def _yc_classify_spec_tag(category, name="", raw_name=""):
    """根据商品分类值和商品名称判断规格和标签.

    匹配顺序（四层优先级）：
      0. 盲盒     — 货品名称含"盲盒"直接归玩具类
      1. toy_exact — 玩具精确词（棋盘/扑克牌等，两字以上，不会和硬分类冲突）
      2. 硬分类    — 水杯 → 餐具(含盘) → 挂件 → 数码 → 箱包
      3. toy_fallback — 玩具泛词兜底
      4. 默认生活类/生活用品

    说明：分类列常是系列名（如"萌宠系列""卡皮巴拉系列"），无法反映货品功能，
    因此硬分类同时匹配商品名称。name 传清洗后的展示名（已去掉 (N/箱) 等
    包装说明），避免"箱/包"等包装词误命中箱包。
    """
    cat = str(category).strip()
    nm = str(name or "").strip()
    text = (cat + " " + nm).strip()

    # --- 第 0 层：货品名称含"盲盒"直接归玩具类（原始名或展示名命中均可）---
    if "盲盒" in nm or "盲盒" in str(raw_name or ""):
        return "玩具类", "儿童玩具"

    # --- 第 1 层：玩具精确词（两字以上完整词，不会和硬分类冲突）---
    toy_exact = ["棋盘", "象棋", "飞行棋", "斗兽棋", "军棋", "围棋", "跳棋",
                 "五子棋", "大富翁", "扑克牌", "卡牌", "纸牌"]
    for kw in toy_exact:
        if kw in cat:
            return "玩具类", "儿童玩具"

    # --- 第 2 层：硬分类 ---
    cup_kws = ["陶瓷杯", "玻璃杯", "马克杯", "保温杯", "水杯", "杯"]
    tableware_kws = ["陶瓷碗", "餐具", "泡面碗", "碗盘", "碗", "盘", "碟", "筷", "勺"]
    pendant_kws = ["公仔扣", "钥匙扣", "挂件", "挂饰"]
    digital_kws = ["数码", "充电", "耳机", "音箱", "数据线"]
    bag_kws = ["时尚小包", "手提包", "包包", "皮包", "包", "袋", "箱"]
    decor_kws = ["水晶球"]

    for kw in cup_kws:
        if kw in text:
            return "生活类", "水杯"
    for kw in tableware_kws:
        if kw in text:
            return "生活类", "餐具"
    for kw in pendant_kws:
        if kw in text:
            return "生活类", "挂件"
    for kw in digital_kws:
        if kw in text:
            return "生活类", "数码"
    for kw in decor_kws:
        if kw in text:
            return "生活类", "摆件"
    for kw in bag_kws:
        if kw in text:
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


def _fmt_amount(v):
    """把合计金额格式化成文件名字符串：整数不带小数点，小数最多保留2位."""
    if v is None:
        return ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v).strip()
    if f == int(f):
        return str(int(f))
    return ("%.2f" % f).rstrip("0").rstrip(".")


def _process_yc_order(parsed, fp, out_dir, multiple=False):
    """把一个订单解析结果写成标准表，返回结果摘要；无数据返回 None."""
    cust = parsed["客户名"] or os.path.splitext(os.path.basename(fp))[0]
    df = parsed["df"]
    img_map = parsed["img_map"]

    print(f"  {os.path.basename(fp)} → 客户={cust!r}, 商品行={len(df)}, 图片={len(img_map)}")

    if len(df) == 0:
        print(f"    [跳过] 无商品数据")
        return None

    rows_data = []
    for pd_idx, row in df.iterrows():
        name = str(row.get("商品名称", "")).strip()
        if not name or name == "nan":
            continue

        category = str(row.get("商品分类", "")).strip()

        # 数量、单价、金额（价格以原表格"单价"列为准）
        qty_v = row.get("数量")
        price_v = row.get("单价")
        try:
            if pd.isna(price_v) or str(price_v).strip() in ("", "nan", "None"):
                price_v = row.get("折后单价")   # 兜底：单价为空时退回折后单价
        except (TypeError, ValueError):
            pass
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
        spec_val, tag_val = _yc_classify_spec_tag(category, display_name, raw_name=name)

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
        return None

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

    # 输出名 = 厂家 + 原表格名称 + 原表格合计金额
    orig_stem = os.path.splitext(os.path.basename(fp))[0]
    amount_str = _fmt_amount(parsed.get("合计金额"))
    if multiple:
        # 同一文件含多个客户时，文件名追加客户名避免互相覆盖
        safe_cust = re.sub(r"[\\/:*?\"<>|\n\r\t]+", "_", cust).strip()
        out_name = f"原创{orig_stem}_{safe_cust}{amount_str}.xlsx"
    else:
        out_name = f"原创{orig_stem}{amount_str}.xlsx"
    out_name = re.sub(r"[\\/:*?\"<>|\n\r\t]+", "_", out_name)
    out_path = os.path.join(out_dir, out_name)

    from excel_processor import _write_output_with_images, _write_ycjh_if_needed

    row_to_header_row = {}
    for out_i, pd_i in enumerate(orig_indices):
        row_to_header_row[out_i] = pd_i + 5

    _write_output_with_images(out_df, img_map, row_to_header_row, out_path)

    # 油菜花：原创管线也处理刘斌/胡欣茹
    _write_ycjh_if_needed(out_df, img_map, row_to_header_row, "原创", cust, orig_stem, out_dir)

    decomp_count = sum(1 for rd in rows_data if rd[0]["_decomposed"])
    img_count = len([v for v in row_to_header_row.values() if v in img_map])
    print(f"  ✓ {out_name}: {len(out_df)} 行（拆解 {decomp_count} 行, 图片 {img_count}）")
    return {"name": out_name, "customer": cust, "rows": len(out_df),
            "spec_changes": decomp_count, "image_count": img_count}


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
        orders = _parse_yc_file(fp)
        if not orders:
            continue
        multiple = len(orders) > 1
        for parsed in orders:
            r = _process_yc_order(parsed, fp, out_dir, multiple=multiple)
            if r:
                results.append(r)

    return results


if __name__ == "__main__":
    process_yuanchuang_all()
