import os
import re
import io
import glob
import tempfile
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter, column_index_from_string
from config import (
    INPUT_DIR, OUTPUT_DIR, TARGET_COLUMNS, COLUMN_MAPPING,
    SPEC_REGEX, SKIP_SPEC_KEYWORDS, FIXED_INTEGRAL_MULTIPLIER,
    UNIT_FIXED_VALUE, DEFAULT_FIELD_VALUES,
)


def _round_half_up(value, ndigits=2):
    """四舍五入（避开 Python round 的银行家舍入）。"""
    import math
    factor = 10 ** ndigits
    return math.floor(value * factor + 0.5) / factor


def find_excel_files(directory):
    """扫描输入目录下的厂家子文件夹，返回 [(厂家名, 文件绝对路径), ...]。

    目录结构：directory/{厂家名}/{表格文件}
    若文件直接放在 directory 下（无子文件夹），厂家名记为"未分类"。
    """
    patterns = ["*.xlsx", "*.xls", "*.csv"]
    results = []

    if not os.path.isdir(directory):
        return results

    # 1) 厂家子文件夹
    for entry in sorted(os.listdir(directory)):
        sub = os.path.join(directory, entry)
        if not os.path.isdir(sub):
            continue
        supplier = entry.strip()
        for ext in patterns:
            for f in sorted(glob.glob(os.path.join(sub, ext))):
                results.append((supplier, f))

    # 2) 直接放在根目录下的文件（无子文件夹）
    for ext in patterns:
        for f in sorted(glob.glob(os.path.join(directory, ext))):
            results.append(("未分类", f))

    return results


# 原始文件名中常见的单据类后缀（提取店铺/收货人时需去掉）
_STORE_NAME_SUFFIXES = [
    "销售单", "销售明细", "销售订单", "销售", "订单", "采购单", "采购订单",
    "送货单", "发货单", "出库单", "清单", "明细表", "对账单", "单据", "报价单",
    "销售报表", "出货单",
]


def _extract_from_address_segment(segment):
    """从"地址+收货人+电话"片段中提取 (短地址, 收货人)。

    例如 "西充县诚信大道158号文艳13330778702" → ("西充", "文艳")
    """
    s = segment
    # 去掉电话号码（连续数字，>=7位）
    s = re.sub(r"\d{7,}", "", s)
    # 去掉地址中的数字部分（门牌号等）
    s = re.sub(r"\d+", "", s)
    # 提取短地址（县/市/区前的2-3个字），并把"XX县/市/区"整体移除
    location = ""
    m = re.search(r"([\u4e00-\u9fff]{2,3})(?:县|市|区)", s)
    if m:
        location = m.group(1)
        s = s[:m.start()] + s[m.end():]
    # 去掉路名+道路类型（如"诚信大道"、"人民南路"）：去掉"大道/路/街"及其前面2-4个汉字
    s = re.sub(r"[\u4e00-\u9fff]{2,4}(?:大道|路|街)", "", s)
    # 去掉其他地址关键词
    addr_keywords = ["省", "号", "幢", "栋", "单元", "楼", "层", "室"]
    for kw in addr_keywords:
        s = s.replace(kw, "")
    # 去掉首尾标点和空白
    s = re.sub(r"^[\s\-_（）()\[\]【】,，。.、]+", "", s)
    s = re.sub(r"[\s\-_（）()\[\]【】,，。.、]+$", "", s)
    recipient = s.strip()
    return location, recipient


def extract_store_name(filename, supplier_name=""):
    """从原始文件名中提取店铺/收货人名称。

    支持的命名格式：
      A) 厂家名-店铺名销售单.xlsx  → 提取"店铺名"
      B) 销售单（销售员）地址收货人电话（物流）单号(序号).xls
         → 提取"短地址+收货人"，如 "西充文艳"

    若无法提取到有效名称，返回空串（调用方只用厂家名命名）。
    """
    name = os.path.splitext(os.path.basename(filename))[0].strip()

    # 去掉厂家名前缀
    if supplier_name and name.startswith(supplier_name):
        name = name[len(supplier_name):]

    # 去掉开头的单据类型前缀（销售单、订单、送货单等）
    for prefix in _STORE_NAME_SUFFIXES:
        if name.startswith(prefix):
            name = name[len(prefix):]
            break

    # 格式 B：销售单（销售员）地址收货人电话（物流）单号(序号)
    # 找到第一个右括号到下一个左括号之间的片段
    m = re.search(r"[）)]\s*([^（）()]+?)\s*[（(]", name)
    if m:
        segment = m.group(1).strip()
        location, recipient = _extract_from_address_segment(segment)
        if recipient:
            return (location + recipient) if location else recipient

    # 格式 C：店名在括号前（如 "都江堰  万丽妹（郑常莉）.xls"）
    # 取第一个左括号前的文本作为店名
    m = re.search(r"^([^（(]+)[（(]", name)
    if m:
        candidate = m.group(1).strip()
        # 清理多余空格
        candidate = re.sub(r"\s+", "", candidate)
        if candidate and len(candidate) >= 2 and not re.fullmatch(r"\d+", candidate):
            return candidate

    # 回退 A：提取第一个括号中的内容作为收货人/店铺名
    m = re.search(r"[（(]([^（）()]{1,20})[）)]", name)
    if m:
        candidate = m.group(1).strip()
        if candidate and not re.fullmatch(r"\d+", candidate):
            return candidate

    # 回退：去掉末尾序号 (2)、单号 XS2026...、电话号码
    name = re.sub(r"\(\d+\)\s*$", "", name)
    name = re.sub(r"[A-Za-z]+\d+\s*$", "", name)
    name = re.sub(r"\d{7,}\s*$", "", name)

    # 反复去掉末尾的单据后缀
    changed = True
    while changed:
        changed = False
        for suffix in _STORE_NAME_SUFFIXES:
            if name.endswith(suffix):
                name = name[: -len(suffix)]
                changed = True
                break

    name = re.sub(r"^[\s\-_]+", "", name)
    name = re.sub(r"[\s\-_]+$", "", name)
    name = name.strip()

    if not name or name == supplier_name:
        return ""
    return name


def _read_images_from_wb(wb):
    """读取工作簿中的图片，按行号映射。同一行多张图时取最左列的（图片1）。"""
    img_rows = {}
    for sheet in wb.worksheets:
        if not hasattr(sheet, "_images"):
            continue
        for img_obj in sheet._images:
            try:
                anchor = img_obj.anchor
                if hasattr(anchor, "_from"):
                    row_idx = anchor._from.row
                    col_idx = anchor._from.col
                elif hasattr(anchor, "row"):
                    row_idx = anchor.row
                    col_idx = getattr(anchor, "col", 0)
                else:
                    continue
                key = row_idx + 1
                # 同行只保留最左列的图片（对应"图片1"列）
                if key not in img_rows or col_idx < img_rows[key][0]:
                    img_rows[key] = (col_idx, img_obj)
            except Exception:
                continue
    # 只返回图片对象，丢弃列索引
    return {k: v[1] for k, v in img_rows.items()}


def _resize_image_bytes(data, max_w, max_h):
    """用 PIL 将图片字节按比例缩放到不超过 max_w×max_h，返回新的 PNG 字节。"""
    try:
        from PIL import Image as PILImage
        pil = PILImage.open(io.BytesIO(data))
        pil = pil.convert("RGB")
        orig_w, orig_h = pil.size
        scale = min(max_w / orig_w, max_h / orig_h)
        new_size = (int(orig_w * scale), int(orig_h * scale))
        pil = pil.resize(new_size, PILImage.LANCZOS)
        out = io.BytesIO()
        pil.save(out, format="JPEG", quality=85)
        return out.getvalue()
    except Exception:
        return data


def _copy_image(img_obj, max_w=None, max_h=None):
    try:
        from openpyxl.drawing.image import Image
        data = img_obj._data()
        resized = False
        if max_w and max_h:
            new_data = _resize_image_bytes(data, max_w, max_h)
            if new_data is not data:
                data = new_data
                resized = True
        new_img = Image(io.BytesIO(data))
        if not resized:
            try:
                new_img.width = img_obj.width
                new_img.height = img_obj.height
            except Exception:
                pass
        return new_img
    except Exception:
        return None


def _find_header_row(wb):
    sheet = wb.active
    max_scan = min(50, sheet.max_row)
    for row_idx in range(1, max_scan + 1):
        row_vals = [str(cell.value).strip() if cell.value is not None else ""
                    for cell in sheet[row_idx]]
        non_empty = [v for v in row_vals if v]
        if len(non_empty) >= 2:
            has_name = False
            for target_aliases in COLUMN_MAPPING.values():
                for v in row_vals:
                    if v in target_aliases:
                        has_name = True
                        break
                if has_name:
                    break
            if has_name:
                return row_idx, row_vals
    return 1, [str(c.value).strip() if c.value is not None else "" for c in sheet[1]]


def _find_header_row_pandas(raw_df):
    """用 pandas 扫描前 50 行查找表头行（用于 .xls 等 openpyxl 不支持的格式）。"""
    max_scan = min(50, len(raw_df))
    for row_idx in range(max_scan):
        row_vals = [str(v).strip() if v is not None else "" for v in raw_df.iloc[row_idx].tolist()]
        non_empty = [v for v in row_vals if v]
        if len(non_empty) >= 2:
            for target_aliases in COLUMN_MAPPING.values():
                if any(v in target_aliases for v in row_vals):
                    return row_idx + 1, row_vals  # 1-based
    return 1, [str(v) for v in raw_df.iloc[0].tolist()]


def _convert_xls_to_xlsx(src_xls, dst_xlsx):
    """通过 Excel COM 将 .xls 转换为 .xlsx（保留图片）。

    返回 True 表示成功，False 表示转换失败（调用方应回退到 xlrd 纯数据读取）。
    """
    import subprocess

    ps_script = (
        "$ErrorActionPreference='Stop';"
        "$excel = New-Object -ComObject Excel.Application;"
        "$excel.Visible = $false;"
        "$excel.DisplayAlerts = $false;"
        f"$wb = $excel.Workbooks.Open('{src_xls}');"
        f"$wb.SaveAs('{dst_xlsx}', 51);"
        "$wb.Close($false);"
        "$excel.Quit();"
        "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null;"
        "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null;"
    )
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", ps_script],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            return False
        return os.path.exists(dst_xlsx)
    except Exception:
        return False


def _finalize_xlsx_package(xlsx_path, sheet_name="Sheet1"):
    """openpyxl 保存后补全 OOXML 部件。

    openpyxl 写出的 docProps/app.xml 缺少 HeadingPairs/TitlesOfParts，
    部分上传平台的解析器从 TitlesOfParts 读取 sheet 名，缺失会报
    "sheet名未获取到"。这里直接在 zip 包内重写 app.xml 补全该部件，
    不经过 Excel/WPS COM 重新保存（WPS 重保存会写入非法 Content_Types
    如 image/.jpg，导致平台拒收整个文件）。
    """
    import zipfile, shutil

    app_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Microsoft Excel</Application>'
        '<DocSecurity>0</DocSecurity>'
        '<ScaleCrop>false</ScaleCrop>'
        '<HeadingPairs><vt:vector size="2" baseType="variant">'
        '<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>'
        '<vt:variant><vt:i4>1</vt:i4></vt:variant>'
        '</vt:vector></HeadingPairs>'
        f'<TitlesOfParts><vt:vector size="1" baseType="lpstr">'
        f'<vt:lpstr>{sheet_name}</vt:lpstr></vt:vector></TitlesOfParts>'
        '<Company></Company>'
        '<LinksUpToDate>false</LinksUpToDate>'
        '<SharedDoc>false</SharedDoc>'
        '<HyperlinksChanged>false</HyperlinksChanged>'
        '<AppVersion>16.0300</AppVersion>'
        '</Properties>'
    )

    tmp_path = xlsx_path + ".tmp"
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin, \
             zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "docProps/app.xml":
                    zout.writestr(item, app_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))
        shutil.move(tmp_path, xlsx_path)
        return True
    except Exception as e:
        print(f"    [警告] xlsx 补全失败: {e}")
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        return False


def _sanitize_content_types(xlsx_path):
    """修正 xlsx 包内 [Content_Types].xml 的非法条目。

    WPS 重保存时可能写入 `<Default Extension="JPG" ContentType="image/.jpg"/>`
    这类非法内容类型，严格的平台解析器会拒收整个文件。这里按扩展名统一
    改为标准 MIME 类型。
    """
    import zipfile, shutil, re

    ext_mime = {
        "jpg": "image/jpeg", "jpeg": "image/jpeg", "jpe": "image/jpeg",
        "png": "image/png", "gif": "image/gif", "bmp": "image/bmp",
        "tif": "image/tiff", "tiff": "image/tiff",
        "emf": "image/x-emf", "wmf": "image/x-wmf",
    }
    valid_img_mime = set(ext_mime.values()) | {"image/x-icon"}

    tmp_path = xlsx_path + ".tmp"
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin:
            ct = zin.read("[Content_Types].xml").decode("utf-8")
            fixed = ct

            def _fix_default(m):
                ext, mime = m.group(1), m.group(2)
                if mime.startswith("image/") and mime not in valid_img_mime:
                    correct = ext_mime.get(ext.lower())
                    if correct:
                        return f'<Default Extension="{ext}" ContentType="{correct}"/>'
                return m.group(0)

            fixed = re.sub(
                r'<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"\s*/>',
                _fix_default, fixed,
            )

            if fixed == ct:
                return True

            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "[Content_Types].xml":
                        zout.writestr(item, fixed)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        shutil.move(tmp_path, xlsx_path)
        return True
    except Exception as e:
        print(f"    [警告] Content_Types 修正失败: {e}")
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        return False


def _insert_images_with_com(xlsx_path, sheet_name, image_items):
    """用 Excel/WPS COM 把图片以 AddPicture 方式（等效复制粘贴）插入单元格。

    image_items: [(row_1based, col_1based, jpeg_bytes), ...]
    图片位置=目标单元格的 Left/Top/Width/Height，Placement=1（随单元格
    改变大小和位置）。返回 True 表示成功。
    """
    import subprocess, json, tempfile

    tmp_dir = os.path.join(tempfile.gettempdir(), "_com_imgs_" + str(os.getpid()))
    os.makedirs(tmp_dir, exist_ok=True)
    manifest = {"file": xlsx_path, "sheet": sheet_name, "images": []}
    try:
        for idx, (r, c, data) in enumerate(image_items):
            img_path = os.path.join(tmp_dir, f"img_{idx}.jpg")
            with open(img_path, "wb") as fp:
                fp.write(data)
            manifest["images"].append({"path": img_path, "row": r, "col": c})

        manifest_path = os.path.join(tmp_dir, "manifest.json")
        with open(manifest_path, "w", encoding="utf-8") as fp:
            json.dump(manifest, fp, ensure_ascii=False)

        ps_script = (
            "$ErrorActionPreference='Stop';"
            f"$m = Get-Content -Raw -Path '{manifest_path}' -Encoding UTF8 | ConvertFrom-Json;"
            "$excel = New-Object -ComObject Excel.Application;"
            "$excel.Visible = $false;"
            "$excel.DisplayAlerts = $false;"
            "$wb = $excel.Workbooks.Open($m.file);"
            "$ws = $wb.Worksheets.Item($m.sheet);"
            "foreach ($img in $m.images) {"
            "  $cell = $ws.Cells.Item([int]$img.row, [int]$img.col);"
            "  $imgSize = 80;"
            "  $shape = $ws.Shapes.AddPicture($img.path, $false, $true,"
            "     $cell.Left, $cell.Top, $imgSize, $imgSize);"
            "  $shape.Placement = 1;"
            "  $shape.Left = $cell.Left + [Math]::Round(($cell.Width - $imgSize) / 2);"
            "  $shape.Top = $cell.Top + [Math]::Round(($cell.Height - $imgSize) / 2);"
            "}"
            "$wb.Save();"
            "$wb.Close($true);"
            "$excel.Quit();"
            "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null;"
            "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null;"
            "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null;"
        )
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", ps_script],
            capture_output=True, text=True, timeout=180,
        )
        if result.returncode != 0:
            print(f"    [警告] COM 插图脚本失败: {(result.stderr or result.stdout)[:300]}")
            return False
        print(f"    🖼 COM 插入图片 {len(image_items)} 张")
        return True
    except Exception as e:
        print(f"    [警告] COM 插图异常: {e}")
        return False
    finally:
        try:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass


def read_with_images(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(filepath)
        return df, {}, None

    # .xls 旧格式：openpyxl 不支持。先尝试用 Excel COM 转成 .xlsx 以保留图片，
    # 转换失败再回退到 pandas+xlrd 纯数据读取（无图片）。
    if ext == ".xls":
        import tempfile

        tmp_xlsx = os.path.join(tempfile.gettempdir(), "_xls_conv_" + str(os.getpid()) + ".xlsx")
        converted = _convert_xls_to_xlsx(os.path.abspath(filepath), tmp_xlsx)
        if converted:
            try:
                wb = load_workbook(tmp_xlsx, data_only=True)
                header_row_idx, headers = _find_header_row(wb)
                img_map = _read_images_from_wb(wb)
                df = pd.read_excel(tmp_xlsx, header=header_row_idx - 1, dtype=object)
                df.columns = [str(c).strip() for c in df.columns]
                df = df.dropna(how="all")
                row_to_header_row = {}
                for row_idx in range(header_row_idx + 1, header_row_idx + 1 + len(df)):
                    data_idx = row_idx - header_row_idx - 1
                    row_to_header_row[data_idx] = row_idx
                wb.close()
                return df, img_map, row_to_header_row
            finally:
                try:
                    os.remove(tmp_xlsx)
                except OSError:
                    pass

        # 回退：xlrd 纯数据读取（无图片）
        raw = pd.read_excel(filepath, header=None, dtype=object, engine="xlrd")
        header_row_idx, headers = _find_header_row_pandas(raw)
        df = raw.iloc[header_row_idx:].copy()
        df.columns = headers
        df = df.dropna(how="all").reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        row_to_header_row = {}
        for data_idx in range(len(df)):
            row_to_header_row[data_idx] = header_row_idx + 1 + data_idx
        return df, {}, row_to_header_row

    wb = load_workbook(filepath, data_only=True)
    header_row_idx, headers = _find_header_row(wb)
    img_map = _read_images_from_wb(wb)

    df = pd.read_excel(filepath, header=header_row_idx - 1, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    row_to_header_row = {}
    for row_idx in range(header_row_idx + 1, header_row_idx + 1 + len(df)):
        data_idx = row_idx - header_row_idx - 1
        row_to_header_row[data_idx] = row_idx

    wb.close()
    return df, img_map, row_to_header_row


def clean_dataframe(df):
    df = df.copy()
    df = df.dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.replace('\n', ' ').str.strip()
        df[col] = df[col].replace({"nan": "", "None": ""})
    return df


def map_columns(df, supplier_name=""):
    source_cols = {c.lower(): c for c in df.columns}
    rename_map = {}

    for target, aliases in COLUMN_MAPPING.items():
        if target == "展示名":
            continue
        for alias in aliases:
            if alias.lower() in source_cols:
                rename_map[source_cols[alias.lower()]] = target
                break

    # 小小玩具：原表格"型号"列对应目标"尺寸"列（而非"规格"）
    if supplier_name == "小小玩具":
        for c in df.columns:
            if str(c).strip().lower() == "型号":
                rename_map[c] = "尺寸"
                break

    # 保留"备注"列供规格拆解使用（如"盒(6个)"表示倍数），不参与输出
    remark_col = None
    for c in df.columns:
        if str(c).strip() in ("备注", "备注栏", "说明", "remarks", "Remarks"):
            remark_col = c
            break

    df = df.rename(columns=rename_map)
    if remark_col:
        df["_备注"] = df[remark_col]

    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    keep = list(TARGET_COLUMNS)
    if "_备注" in df.columns:
        keep.append("_备注")
    df = df[keep]

    mapped = [c for c in rename_map.values()]
    unmatched = [c for c in TARGET_COLUMNS if c != "展示名" and c not in mapped]
    return df, rename_map, unmatched


def _has_skip_keyword(name):
    for kw in SKIP_SPEC_KEYWORDS:
        if kw in name:
            return True
    return False


def _to_num(s):
    """把字符串转成数字，整数则返回 int，否则 float。"""
    v = float(s)
    return int(v) if v.is_integer() else v


def _extract_spec_from_text(text):
    """从文本中提取规格信息。

    返回 (multiplier, spec_parts, cleaned_text)：
      - multiplier: 数量换算倍数（乘法表达式与数量单位的乘积）
      - spec_parts: 匹配到的规格文本片段列表
      - cleaned_text: 移除规格片段后的残留文本
    体积/重量单位仅进入 spec_parts，不参与 multiplier 计算。
    """
    multiplier = 1
    spec_parts = []
    cleaned = text
    count_applied = False  # 只取第一个 count 匹配作为倍数

    for m in re.finditer(SPEC_REGEX, text, re.IGNORECASE):
        matched = m.group(0)

        if m.group("count") and count_applied:
            # 已拆解过一次，后续 count（如"8瓶"、"4瓶"）不再处理：不删、不加 spec_parts
            continue

        spec_parts.append(matched)
        # 只替换当前这一处匹配，避免同名片段被误删
        cleaned = cleaned.replace(matched, "", 1)

        if m.group("mult"):
            a = _to_num(m.group("mult_a"))
            b = _to_num(m.group("mult_b"))
            mult_unit = m.group("mult_unit")
            unit_val = UNIT_FIXED_VALUE.get(mult_unit, 1) if mult_unit else 1
            multiplier *= a * b * unit_val
        elif m.group("add"):
            # 加法表达式（赠品）：6+1入 → 倍数取 6，+1 为赠品不计入
            a = _to_num(m.group("add_a"))
            add_unit = m.group("add_unit")
            unit_val = UNIT_FIXED_VALUE.get(add_unit, 1) if add_unit else 1
            multiplier *= a * unit_val
        elif m.group("count"):
            # 只取第一个 count 作为倍数（如"24入3入"只取24，不乘3）
            num = _to_num(m.group("count_num"))
            unit = m.group("count_unit")
            multiplier *= num * UNIT_FIXED_VALUE.get(unit, 1)
            count_applied = True
        # vw（体积/重量）：只记录规格，不改变倍数

    return multiplier, spec_parts, cleaned


# 已知需要从展示名中移除的注释类前缀（如"不退换"、"整盒"、"单个"）
_NAME_NOTE_PREFIXES = ["不退换", "不退货", "不退", "特价", "清仓", "整盒", "单个", "单包", "活动款"]


def _clean_display_name(name):
    """清理拆解后的展示名。

    移除：
      - 前缀产品编码/注释：如 "(2121413)"、"(ABC-46)"、"(不退换)"、"不退换(SP-2RYZB)"
      - 全角括号内容：如 "（蘑菇兔）"
      - 价格：如 "1元"、"(5元)"、"(2元)"
      - 后缀编码：如 "(8833)"
      - 中文数字规格片段：如 "五入"、"十入"
      - 残留的"装/裝"、连接符、标点、多余括号与空白
    """
    # 半角/全角括号统一处理
    # 1) 去掉开头的注释词（不退换等），无论是否带括号，并去掉紧随其后的编码括号
    for note in _NAME_NOTE_PREFIXES:
        # 形如 "不退换(SP-2RYZB)" 或 "(不退换)(T-2110)"
        name = re.sub(r"^[\s（(]*" + re.escape(note) + r"[\s）)]*", "", name)

    # 2) 反复去掉开头的括号组（编码或空），兼容半角/全角/混合括号
    changed = True
    while changed:
        changed = False
        # 匹配 (code) / （code） / （code) / (code） 等任意组合
        m = re.match(r"^\s*[（(]([A-Za-z0-9\-\.]*)[）)]\s*", name)
        if m:
            name = name[m.end():]
            changed = True

    # 2.5) 去掉开头括号包裹的中文品牌/系列名（如"（蘑菇兔）"、"(淘米文创)"）
    #      只移除括号，保留其中文内容作为展示名的一部分
    m = re.match(r"^\s*[（(]([^（）()]+)[）)]\s*", name)
    if m:
        inner = m.group(1).strip()
        if re.search(r"[\u4e00-\u9fff]", inner) and not re.fullmatch(r"[A-Za-z0-9\-\.]+", inner):
            name = inner + name[m.end():]

    # 3) 去掉价格 "X元" 及价格说明词（如 "零售价49.9元"、"建议零售价"、"零售59.9元"）
    name = re.sub(r"\d+(?:\.\d+)?\s*元", "", name)
    name = re.sub(r"(?:建议)?零售价?", "", name)

    # 4) 去掉末尾括号包裹的编码、价格或规格（如 "(168-48)"、"(2元)"、"(24入)"、"(16件/件)"、"(60/件)"）
    name = re.sub(r"\s*[（(][A-Za-z0-9\-\.]+[）)]\s*$", "", name)
    name = re.sub(r"\s*[（(]\s*\d+(?:\.\d+)?\s*元\s*[）)]\s*$", "", name)
    # 末尾括号内为"数字+规格单位"的规格片段（如"(24入)"、"(6个)"、"(16件/件)"、"(60/件)"、"(192/件16/盒)"、"(108/件 12/盒)"）
    _spec_unit_alt = "|".join(["入", "支", "个", "包", "袋", "盒", "板", "瓶", "罐", "条", "片", "粒", "颗", "张", "本", "件", "中包"])
    # 匹配模式：数字后可选单位，后可跟多组(/数字单位 或 /单位)（如 192/件16/盒、108/件 12/盒、60/件）
    _spec_content = r"\d+(?:\.\d+)?\s*(?:" + _spec_unit_alt + r")?[装裝]?" \
                    r"(?:\s*[/／]?\s*\d*\s*(?:" + _spec_unit_alt + r")?[装裝]?)*"
    name = re.sub(r"\s*[（(]\s*" + _spec_content + r"\s*[）)]\s*$", "", name)

    # 5) 清理残留的空括号 "()" "（）" 及零散的多余括号
    name = re.sub(r"[（(]\s*[）)]", "", name)
    name = re.sub(r"^[）)]+", "", name)
    name = re.sub(r"[（(]+$", "", name)

    # 6) 去掉中文数字+单位 的规格片段（如 "五入"、"十支"）
    name = re.sub(r"[零一二三四五六七八九十百千]+\s*(?:" + "|".join(
        ["入", "支", "个", "包", "袋", "盒", "板", "瓶", "罐", "条", "片", "粒", "颗", "张", "本"]
    ) + r")[装裝]?", "", name)

    # 7) 【重要：在简单单位清理之前】先删完整规格尾部（避免"盒"被 step 7 删掉后"12个入盒"变成"12个入"无法匹配）
    _spec_unit_alt2 = "|".join(["入", "支", "个", "包", "袋", "盒", "板", "瓶", "罐", "条", "片", "粒", "颗", "张", "本", "件", "中包", "大包", "只", "桶", "端"])
    for _ in range(5):
        prev = name
        name = re.sub(r"\s*[（(]\s*\d+(?:\.\d+)?\s*(?:" + _spec_unit_alt2 + r")?\s*[/／、]\s*\d*\s*(?:" + _spec_unit_alt2 + r")?\s*[）)]\s*$", "", name)
        name = re.sub(r"\s*\d+\s*(?:" + _spec_unit_alt2 + r")?\s*入\s*盒\s*$", "", name)
        name = re.sub(r"(\s?)\d+(?:\.\d+)?\s*(?:" + _spec_unit_alt2 + r")?\s*[/／、]\s*(?:" + _spec_unit_alt2 + r")\s*$", r"\1", name)
        if name == prev:
            break

    # 7.1) 中间位置的规格片段：N入（如【拆散】6入酷炫英雄）
    name = re.sub(r"\s*\d+入\s*", "", name)

    # 7.2) 删残留的 N+单位（尾部）和单独数字
    # 注意：不删单独的单位字（如"展示盒"的"盒"是商品名）
    _unit_alt = "|".join(["入", "支", "个", "包", "袋", "盒", "板", "瓶", "罐", "条", "片", "粒", "颗", "张", "本", "只", "件"])
    name = re.sub(r"\s*\d+(?:\.\d+)?\s*(?:" + _unit_alt + r")\s*$", "", name)  # 12只、24个、36盒
    name = re.sub(r"\s+\d+(?:\.\d+)?\s*$", "", name)  # 单独的数字

    # 7.3) 兜底：删残留的 "/件" "/盒" 等
    name = re.sub(r"\s*[/／、]\s*(?:件|盒|包|袋|箱|中包|大包)\s*$", "", name)
    name = re.sub(r"^(?:件|盒|包|袋|箱|中包|大包)\s*", "", name)
    # 不再全局删除"装/裝"字（商品名可能有"换装""拼装"等）
    # 保留 * 号（鸿达产品编码常用）
    name = re.sub(r"\s+", " ", name).strip()
    name = re.sub(r"[，,。.、；;：:]+$", "", name)
    name = re.sub(r"^[，,。.、；;：:]+", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


# ===== 家家乐厂家专属：规格/标签按名称关键词分类 =====
_JJL_TOY_KEYWORDS = [
    "玩具", "益智", "DIY", "diy", "游戏", "积木", "拼图", "陀螺",
    "手工", "趣味", "创意", "潮玩", "盲盒", "手办", "模型", "拼装",
    "遥控", "电动", "毛绒", "娃娃", "公仔", "机器人", "恐龙", "车",
    "枪", "剑", "棒", "圈", "球", "棋", "魔方", "彩绘",
]
_JJL_TAG_CUP = ["保温杯", "陶瓷杯", "玻璃杯", "杯", "壶", "瓶"]
_JJL_TAG_TABLEWARE = ["碗", "筷", "勺", "餐具", "盘", "碟", "刀叉"]
_JJL_TAG_DIGITAL = ["数码", "充电", "耳机", "音箱", "数据线", "充电宝", "U盘", "键盘", "鼠标"]
_JJL_TAG_BAG = ["包", "袋", "箱", "背包", "钱包", "手提袋"]
_JJL_TAG_PENDANT = ["挂件", "挂饰", "挂坠", "钥匙扣"]


def _jjl_classify_spec_tag(name):
    """家家乐：根据货品名称判断规格(玩具类/生活类)和标签。

    优先检查生活类关键词（杯/碗/盘等），避免含"涂鸦"等词的生活用品被误判为玩具。
    扑克/卡牌归为玩具类。
    """
    # 优先匹配生活类关键词
    if any(kw in name for kw in _JJL_TAG_PENDANT):
        return "生活类", "挂件"
    if any(kw in name for kw in _JJL_TAG_CUP):
        return "生活类", "水杯"
    if any(kw in name for kw in _JJL_TAG_TABLEWARE):
        return "生活类", "餐具"
    if any(kw in name for kw in _JJL_TAG_DIGITAL):
        return "生活类", "数码"
    if any(kw in name for kw in _JJL_TAG_BAG):
        return "生活类", "箱包"
    # 再检查玩具类关键词
    if any(kw in name for kw in _JJL_TOY_KEYWORDS):
        return "玩具类", "儿童玩具"
    # 扑克/卡牌归为玩具类
    if "扑克" in name or "卡牌" in name:
        return "玩具类", "儿童玩具"
    return "生活类", "生活用品"


def _extract_deal_amount(filepath):
    """从原始表格的元数据区提取"成交金额"标签右侧单元格的值（整单成交总额）。"""
    try:
        wb = load_workbook(filepath, data_only=True)
    except Exception:
        return None
    try:
        ws = wb.active
        for row in ws.iter_rows(max_row=min(40, ws.max_row)):
            for cell in row:
                val = cell.value
                if val is None:
                    continue
                sval = str(val).strip()
                if "成交金额" in sval:
                    # 取右侧相邻单元格的值
                    next_cell = ws.cell(row=cell.row, column=cell.column + 1)
                    nval = next_cell.value
                    if nval is None:
                        return None
                    try:
                        return float(nval)
                    except (ValueError, TypeError):
                        return nval
        return None
    finally:
        wb.close()


def parse_spec_and_recalc(row, supplier_name="", deal_amount=None):
    original_name = str(row.get("名称", ""))
    spec = str(row.get("规格", ""))
    remark = str(row.get("_备注", ""))
    qty = row.get("进货数量", "")
    price = row.get("单价", "")

    try:
        qty = float(qty) if qty not in ("", None) else 0
    except (ValueError, TypeError):
        qty = 0
    try:
        price = float(price) if price not in ("", None) else 0
    except (ValueError, TypeError):
        price = 0

    is_jjl = supplier_name == "家家乐"

    # ===== 家家乐厂家专属逻辑 =====
    if is_jjl:
        has_zhenghe = "整盒" in original_name
        # 只有"整盒"才拆解，倍数只取名称中的 1*N（如 1*18 表示1盒18个）
        spec_multiple = 1
        cleaned_name = original_name
        if has_zhenghe:
            m = re.search(r"1\s*[*×xX]\s*(\d+(?:\.\d+)?)", original_name)
            if m:
                spec_multiple = _to_num(m.group(1))
                cleaned_name = original_name[:m.start()] + original_name[m.end():]

        display_name = _clean_display_name(cleaned_name)

        # 数量 × 倍数，单价 ÷ 倍数
        if spec_multiple > 1 and qty > 0:
            qty = qty * spec_multiple
            if price > 0:
                price = _round_half_up(price / spec_multiple, 2)

        # 规格/标签按名称关键词分类
        spec_val, tag_val = _jjl_classify_spec_tag(original_name)

        row["名称"] = original_name
        row["展示名"] = display_name if display_name else original_name
        row["规格"] = spec_val
        row["用途"] = DEFAULT_FIELD_VALUES.get("用途", "")
        row["标签"] = tag_val
        row["单位"] = DEFAULT_FIELD_VALUES.get("单位", "个")
        row["进货数量"] = qty if qty != 0 else ""
        row["单价"] = price if price != 0 else ""
        row["积分倍数"] = FIXED_INTEGRAL_MULTIPLIER
        row["_needs_review"] = False
        return row

    # ===== 鸿达厂家专属逻辑 =====
    is_hongda = supplier_name == "鸿达"
    if is_hongda:
        # 鸿达：只对名称含"入"且单价>20且不含"拆散"的进行拆解
        # "/件" "/盒" "/箱" 等不是规格，仅用于展示名清理
        if "拆散" in original_name or price <= 20:
            # 不拆解，但展示名仍需清理（去掉/件等碎片）
            row["名称"] = original_name
            row["展示名"] = _clean_display_name(original_name)
            row["规格"] = DEFAULT_FIELD_VALUES.get("规格", "")
            row["用途"] = DEFAULT_FIELD_VALUES.get("用途", "")
            row["标签"] = DEFAULT_FIELD_VALUES.get("标签", "")
            row["单位"] = DEFAULT_FIELD_VALUES.get("单位", row.get("单位", ""))
            row["进货数量"] = qty if qty != 0 else ""
            row["单价"] = price if price != 0 else ""
            row["积分倍数"] = FIXED_INTEGRAL_MULTIPLIER
            row["_needs_review"] = False
            return row

        # 从名称中提取"N入"倍数
        spec_multiple = 1
        cleaned_name = original_name
        m = re.search(r"(\d+)\s*入", original_name)
        if m:
            spec_multiple = _to_num(m.group(1))
            cleaned_name = original_name[:m.start()] + original_name[m.end():]

        display_name = _clean_display_name(cleaned_name)

        if spec_multiple > 1 and qty > 0:
            qty = qty * spec_multiple
            if price > 0:
                price = _round_half_up(price / spec_multiple, 2)

        row["名称"] = original_name
        row["展示名"] = display_name if display_name else original_name
        row["规格"] = DEFAULT_FIELD_VALUES.get("规格", "")
        row["用途"] = DEFAULT_FIELD_VALUES.get("用途", "")
        row["标签"] = DEFAULT_FIELD_VALUES.get("标签", "")
        row["单位"] = DEFAULT_FIELD_VALUES.get("单位", row.get("单位", ""))
        row["进货数量"] = qty if qty != 0 else ""
        row["单价"] = price if price != 0 else ""
        row["积分倍数"] = FIXED_INTEGRAL_MULTIPLIER
        row["_needs_review"] = False
        return row

    # ===== 小小玩具厂家专属逻辑 =====
    is_xxwj = supplier_name == "小小玩具"
    if is_xxwj:
        # 小小玩具不做规格拆解，按名称关键词/单价/尺寸判定规格和标签
        _xxwj_keywords = ["香袋", "掌中宝", "挂件", "迷你", "桌伴"]
        has_keyword = any(kw in original_name for kw in _xxwj_keywords)

        # 从"尺寸"列（原"型号"列）提取厘米数用于规则判定，但保留原始值输出
        size_raw = str(row.get("尺寸", "")).strip()
        size_num = None
        m = re.search(r"(\d+(?:\.\d+)?)", size_raw)
        if m:
            size_num = float(m.group(1))
        # 尺寸列保留原始值（如"18CM"、"40cmX30cm"）
        row["尺寸"] = size_raw

        # 规则1：名称含关键词 → 芭比（5-10cm）
        # 规则2：不含关键词，但 单价<12 且 尺寸<20cm → 芭比（5-10cm）
        # 规则3：不满足1、2，且 单价≥50 → 超牛（80cm）
        # 规则4：以上都不满足 → 12-18寸（40-50cm）
        if has_keyword:
            spec_val = "芭比（5-10cm）"
        elif price < 12 and size_num is not None and size_num < 20:
            spec_val = "芭比（5-10cm）"
        elif price >= 50:
            spec_val = "超牛（80cm）"
        else:
            spec_val = "12-18寸（40-50cm）"

        row["名称"] = original_name
        row["展示名"] = ""  # 展示名不填
        row["规格"] = spec_val
        row["用途"] = "礼品机,兑换,零售"
        row["标签"] = "公仔"
        row["单位"] = DEFAULT_FIELD_VALUES.get("单位", row.get("单位", ""))
        row["进货数量"] = qty if qty != 0 else ""
        row["单价"] = price if price != 0 else ""
        row["积分倍数"] = FIXED_INTEGRAL_MULTIPLIER
        row["条码"] = ""  # 条码列不填
        row["_needs_review"] = False
        return row

    # ===== 其他厂家（成龙、辉鸿等）通用逻辑 =====
    if _has_skip_keyword(original_name):
        row["规格"] = DEFAULT_FIELD_VALUES.get("规格", spec)
        row["用途"] = DEFAULT_FIELD_VALUES.get("用途", "")
        row["标签"] = DEFAULT_FIELD_VALUES.get("标签", "")
        row["单位"] = DEFAULT_FIELD_VALUES.get("单位", row.get("单位", ""))
        row["进货数量"] = qty if qty != 0 else ""
        row["单价"] = price if price != 0 else ""
        row["积分倍数"] = FIXED_INTEGRAL_MULTIPLIER
        # 跳过规格拆解（故意不拆，无需复核），但展示名仍需清理
        row["_needs_review"] = False
        row["展示名"] = _clean_display_name(original_name)
        return row

    # 1) 先从商品名称提取规格
    mult_from_name, parts_from_name, cleaned_name = _extract_spec_from_text(original_name)

    # 2) 再从"规格"列提取（用于补充规格文本；倍数仅在名称未提供时才采用）
    mult_from_spec = 1
    parts_from_spec = []
    if spec:
        mult_from_spec, parts_from_spec, _ = _extract_spec_from_text(spec)

    # 3) 从"备注"列提取（如"盒(6个)"表示倍数），同样仅在名称未提供时采用
    mult_from_remark = 1
    if remark:
        mult_from_remark, _, _ = _extract_spec_from_text(remark)

    # 名称有倍数就用名称的，否则依次退回规格列、备注列的倍数
    if mult_from_name > 1:
        spec_multiple = mult_from_name
    elif mult_from_spec > 1:
        spec_multiple = mult_from_spec
    else:
        spec_multiple = mult_from_remark

    # 清理展示名
    display_name = _clean_display_name(cleaned_name)

    # 检测"含规格单位但无法确定倍数"的货品，需人工复核
    # （名称含"入"字，但名称、规格、备注都没提取到有效倍数，
    #   如"(0758-49）入超变机甲板"——"入"前不是数字，无法确定多少入）
    needs_review = (
        spec_multiple == 1
        and "入" in original_name
    )
    row["_needs_review"] = needs_review
    if needs_review and display_name:
        display_name = "⚠" + display_name

    # 数量 × 倍数，单价 ÷ 倍数（四舍五入到2位小数）
    if spec_multiple > 1 and qty > 0:
        qty = qty * spec_multiple
        if price > 0:
            price = _round_half_up(price / spec_multiple, 2)

    row["名称"] = original_name
    row["展示名"] = display_name if display_name else original_name
    # 规格列写入平台要求的类别（非拆解出的规格文本，规格仅用于数量/单价换算）
    row["规格"] = DEFAULT_FIELD_VALUES.get("规格", "")
    row["用途"] = DEFAULT_FIELD_VALUES.get("用途", "")
    row["标签"] = DEFAULT_FIELD_VALUES.get("标签", "")
    row["单位"] = DEFAULT_FIELD_VALUES.get("单位", row.get("单位", ""))
    row["进货数量"] = qty if qty != 0 else ""
    row["单价"] = price if price != 0 else ""
    row["积分倍数"] = FIXED_INTEGRAL_MULTIPLIER

    return row


def convert_types(df):
    for col in df.columns:
        df[col] = df[col].fillna("")
    for col in ["进货数量", "单价", "总金额"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: x if pd.notna(x) else "")
    return df


# 合并时保留首条非空值的列
_DEDUP_FIRST_COLS = {
    "名称", "展示名", "条码", "规格", "用途", "标签", "单位",
    "单价", "积分倍数", "供应商商品编码", "尺寸",
}
# 合并时求和的列
_DEDUP_SUM_COLS = {"进货数量", "总金额"}


def deduplicate_products(df, img_map, row_to_header_row):
    """合并重复货品。

    去重键：条码（strip 后非空）优先，否则用展示名。
    合并规则：
      - 进货数量：求和
      - 单价/名称/展示名/条码/规格/用途/标签/单位/积分倍数/供应商商品编码/尺寸：保留首条非空
      - 图片：取组内首张有图的原始行
    """
    if len(df) <= 1:
        return df, row_to_header_row

    df = df.reset_index(drop=True)

    def _dedup_key(row):
        # 用原始名称（含产品编码）而非展示名，避免不同编码的同名产品被误合并
        name = str(row.get("名称", "")).strip()
        price = str(row.get("单价", "")).strip()
        return f"{name}||{price}"

    keys = df.apply(_dedup_key, axis=1).tolist()

    # 按出现顺序分组
    order = []
    groups = {}
    for i, k in enumerate(keys):
        if k not in groups:
            groups[k] = []
            order.append(k)
        groups[k].append(i)

    if len(order) == len(df):
        # 无重复
        return df, row_to_header_row

    merged_rows = []
    new_row_to_header = {}
    has_img = bool(img_map)

    for out_idx, k in enumerate(order):
        indices = groups[k]
        group_df = df.iloc[indices]
        merged = {}

        # 首条非空列
        for col in _DEDUP_FIRST_COLS:
            if col in group_df.columns:
                vals = group_df[col].tolist()
                val = next(
                    (v for v in vals if v not in ("", None) and not (isinstance(v, float) and pd.isna(v))),
                    vals[0],
                )
                merged[col] = val

        # 求和列
        for col in _DEDUP_SUM_COLS:
            if col in group_df.columns:
                nums = pd.to_numeric(group_df[col], errors="coerce")
                total = nums.sum()
                merged[col] = total if not pd.isna(total) else ""

        # 其余列：取首条
        for col in group_df.columns:
            if col not in merged:
                merged[col] = group_df[col].iloc[0]

        merged_rows.append(merged)

        # 图片映射：组内首张有图的原始行；都没图则取组内首行
        chosen_orig = None
        for i in indices:
            orig = row_to_header_row.get(i) if row_to_header_row else None
            if orig is not None and has_img and orig in img_map:
                chosen_orig = orig
                break
        if chosen_orig is None:
            first_i = indices[0]
            chosen_orig = row_to_header_row.get(first_i) if row_to_header_row else None
        new_row_to_header[out_idx] = chosen_orig

    new_df = pd.DataFrame(merged_rows, columns=df.columns)
    dup_count = len(df) - len(new_df)
    if dup_count > 0:
        print(f"    🔁 合并重复货品: {len(df)} → {len(new_df)} 条（减少 {dup_count} 条）")
    return new_df, new_row_to_header


def _write_output_with_images(df, img_map, row_to_header_row, output_path):
    """写入标准表格，并把源表格自带图片填入"图片"列。

    图片采用"填充于单元格内"策略：设定固定的单元格尺寸，将图片按比例缩放
    后填入，而不是用图片尺寸去撑大单元格。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for col_idx, col_name in enumerate(TARGET_COLUMNS, 1):
        ws.cell(row=1, column=col_idx, value=col_name)

    img_col_idx = TARGET_COLUMNS.index("图片") + 1 if "图片" in TARGET_COLUMNS else None

    # 图片列与单元格尺寸（Excel 列宽单位≈字符宽，行高单位=磅；1磅≈1.333像素）
    IMG_COL_WIDTH = 14          # ≈ 100px
    IMG_ROW_HEIGHT = 100        # 100磅 ≈ 133px
    CELL_W_PX = 100
    CELL_H_PX = 100
    # 图片实际分辨率（高于显示尺寸，保证清晰度）
    IMG_RES_W = 400
    IMG_RES_H = 400

    if img_col_idx:
        ws.column_dimensions[get_column_letter(img_col_idx)].width = IMG_COL_WIDTH

    # 收集待插入图片：(输出行号1基, 列号1基, 图片JPEG字节)
    # 图片不通过 openpyxl 写入（openpyxl 生成的 drawing 结构部分平台解析器
    # 不识别），改为保存后用 Excel/WPS COM 的 AddPicture 方式插入（等效
    # 复制粘贴），Placement=随单元格改变大小和位置。
    image_items = []

    for out_row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        for col_idx, col_name in enumerate(TARGET_COLUMNS, 1):
            if col_name == "图片":
                continue
            ws.cell(row=out_row_idx, column=col_idx, value=row.get(col_name, ""))

        # 固定行高（图片不撑大单元格）
        ws.row_dimensions[out_row_idx].height = IMG_ROW_HEIGHT

        img_data = None

        # 用源表格自带的图片（按高分辨率缩放，保证清晰度）
        if img_map and row_to_header_row and img_col_idx:
            orig_row_num = row_to_header_row.get(out_row_idx - 2)
            if orig_row_num is not None and orig_row_num in img_map:
                try:
                    raw = img_map[orig_row_num]._data()
                    img_data = _resize_image_bytes(raw, IMG_RES_W, IMG_RES_H)
                except Exception:
                    img_data = None

        if img_data and img_col_idx:
            image_items.append((out_row_idx, img_col_idx, img_data))

    # 列宽自适应：按每列最大内容长度计算宽度（图片列保持固定宽度）
    for col_idx, col_name in enumerate(TARGET_COLUMNS, 1):
        if img_col_idx and col_idx == img_col_idx:
            continue
        max_len = len(str(col_name))
        for r in range(2, ws.max_row + 1):
            val = ws.cell(row=r, column=col_idx).value
            if val is not None:
                # 中文字符按2个宽度计算，数字/英文按1个
                s = str(val)
                cnt = sum(2 if ord(c) > 127 else 1 for c in s)
                if cnt > max_len:
                    max_len = cnt
        width = min(max(max_len + 2, 8), 50)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(output_path)
    abs_path = os.path.abspath(output_path)

    if image_items:
        # 用 COM 复制粘贴方式插图，保存后修正 WPS 可能写入的非法 Content_Types
        ok = _insert_images_with_com(abs_path, ws.title, image_items)
        if not ok:
            print("    [警告] COM 插图失败，输出文件将不含图片")
        else:
            _sanitize_content_types(abs_path)
    else:
        # 无图片：直接补全 docProps/app.xml 的 HeadingPairs/TitlesOfParts
        _finalize_xlsx_package(abs_path, sheet_name=ws.title)


def process_single_file(filepath, supplier_name=""):
    filename = os.path.basename(filepath)
    print(f"\n  文件: {filename}")
    if supplier_name:
        print(f"    厂家: {supplier_name}")

    df, img_map, row_to_header_row = read_with_images(filepath)
    original_rows = len(df)

    # 提取整单"成交金额"（用于小小玩具的总金额列）
    deal_amount = _extract_deal_amount(filepath) if supplier_name == "小小玩具" else None
    if supplier_name == "小小玩具":
        print(f"    成交金额: {deal_amount}")

    df = clean_dataframe(df)
    df, rename_map, unmatched = map_columns(df, supplier_name=supplier_name)

    print(f"    原始表头: {list(df.columns)}")
    print(f"    列名映射: {rename_map if rename_map else '(无匹配)'}")
    if unmatched:
        print(f"    ⚠ 未匹配到的目标列: {unmatched}")
    if img_map:
        print(f"    🖼 检测到 {len(img_map)} 张图片")

    spec_change_count = 0
    skip_count = 0
    review_items = []  # 需人工复核的货品名
    def track_spec(row):
        nonlocal spec_change_count, skip_count
        name = str(row.get("名称", ""))
        if _has_skip_keyword(name):
            skip_count += 1
        elif re.search(SPEC_REGEX, name):
            spec_change_count += 1
        result = parse_spec_and_recalc(row, supplier_name=supplier_name, deal_amount=deal_amount)
        if result.get("_needs_review"):
            review_items.append(str(result.get("名称", "")))
        return result

    df = df.apply(track_spec, axis=1)
    # 备注列、复核标记列仅用于处理过程，不输出
    df = df.drop(columns=["_备注", "_needs_review"], errors="ignore")
    df = convert_types(df)

    # 合并重复货品（按条码，无条码按展示名）
    df, row_to_header_row = deduplicate_products(df, img_map, row_to_header_row)

    # 过滤非数据行：合计行、欠款行、表头残留行、备注行等
    _non_data_keywords = ["合计", "欠款", "注：", "注:", "本次成交", "上次欠款"]
    _header_residue = {"商品全名", "行号", "名称", "货品名称"}  # 表头行残留
    def _is_data_row(row):
        name = str(row.get("名称", "")).strip()
        if not name:
            return False
        if name in _header_residue:
            return False
        for kw in _non_data_keywords:
            if kw in name:
                return False
        # 过滤掉无条码且无数量且无单价的行
        barcode = str(row.get("条码", "")).strip()
        qty = str(row.get("进货数量", "")).strip()
        price = str(row.get("单价", "")).strip()
        if not barcode and (not qty or qty == "0" or qty == "0.0") and (not price or price == "nan"):
            return False
        return True

    df = df[df.apply(_is_data_row, axis=1)]

    # 小小玩具：总金额取整单"成交金额"，仅写入第2行（首条数据行），其余行清空
    if supplier_name == "小小玩具" and deal_amount is not None:
        df["总金额"] = ""
        if len(df) > 0:
            df.loc[df.index[0], "总金额"] = deal_amount

    # 输出目录：OUTPUT_DIR/{厂家名}/  文件名：使用原表格名称
    store_name = extract_store_name(filename, supplier_name)
    safe_supplier = supplier_name if supplier_name else "未分类"
    out_dir = os.path.join(OUTPUT_DIR, safe_supplier)
    os.makedirs(out_dir, exist_ok=True)
    # 使用原文件名（去掉扩展名），文件名=厂家+原名，统一 .xlsx
    orig_stem = os.path.splitext(os.path.basename(filename))[0]
    out_basename = f"{safe_supplier}{orig_stem}.xlsx"
    # 文件名中的非法字符替换
    out_basename = re.sub(r'[\\/:*?"<>|]', "_", out_basename)
    output_path = os.path.join(out_dir, out_basename)

    _write_output_with_images(df, img_map, row_to_header_row, output_path)

    print(f"    ✓ 输出: {out_basename}  ({original_rows} → {len(df)} 条)")
    if spec_change_count > 0:
        print(f"    🔧 规格拆解: {spec_change_count} 行")
    if skip_count > 0:
        print(f"    ⏭ 跳过拆解(含拆散等标记): {skip_count} 行")
    if review_items:
        print(f"    ⚠ 需人工复核（含规格单位但无法确定倍数，已用原数量单价，展示名带⚠）: {len(review_items)} 行")
        for nm in review_items:
            print(f"        - {nm}")

    return {
        "name": out_basename,
        "supplier": safe_supplier,
        "input": filepath,
        "output": output_path,
        "rows": len(df),
        "spec_changes": spec_change_count,
        "skip_count": skip_count,
        "review_items": review_items,
        "unmatched_cols": unmatched,
        "image_count": len(img_map) if img_map else 0,
    }


def process_all():
    print("=" * 60)
    print("步骤 1: 扫描并处理所有订单表格")
    print("=" * 60)

    files = find_excel_files(INPUT_DIR)
    if not files:
        print(f"[警告] 未在 {INPUT_DIR} 找到任何表格文件")
        print(f"  请在该目录下按厂家建子文件夹，并放入厂家给的原始表格")
        return []

    print(f"  输入目录: {INPUT_DIR}")
    print(f"  输出目录: {OUTPUT_DIR}")
    print(f"  共发现 {len(files)} 个表格文件")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 处理前先清空输出目录中的旧文件，避免残留
    import shutil
    for entry in os.listdir(OUTPUT_DIR):
        entry_path = os.path.join(OUTPUT_DIR, entry)
        if os.path.isdir(entry_path):
            shutil.rmtree(entry_path, ignore_errors=True)
        elif os.path.isfile(entry_path):
            try:
                os.remove(entry_path)
            except Exception:
                pass

    results = []
    for supplier_name, f in files:
        try:
            result = process_single_file(f, supplier_name=supplier_name)
            results.append(result)
        except Exception as e:
            import traceback
            print(f"    ✗ 失败 [{supplier_name}]: {e}")
            traceback.print_exc()

    print()
    print(f"[完成] 成功处理 {len(results)}/{len(files)} 个文件")
    return results


if __name__ == "__main__":
    process_all()
