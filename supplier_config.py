# -*- coding: utf-8 -*-
"""
供应商配置模块
==============
把「开票状态 / 折扣 / 收款信息」从脚本硬编码抽离为外部 JSON 配置文件（supplier_config.json），
GUI 负责读写，拆分引擎负责读取。

配置结构：
{
  "suppliers": {
    "成龙": {
      "invoice_status": "开票不加点",   # 三档：不开票 / 开票不加点 / 开票加一个点
      "discount": 1.00,                # 折扣系数（1.0=无折扣，0.99=99折，0.98=98折）
      "bank": "……",                    # 开票收款信息（多行文本）
      "no_bank": "……"                  # 不开票收款信息（多行文本）
    }
  }
}

首次运行时若配置文件不存在，则从旧「按店铺对账」文件的「汇总」Sheet 扫描所有供应商的
收款信息 / 不开票收款信息 / 备注，结合内置种子规则（SEED_RULES / SEED_NO_INVOICE）
生成初始配置并落盘。
"""
import json
import sys
from pathlib import Path

# 内置种子（仅用于首次初始化；此后以 supplier_config.json 为准）
SEED_RULES = {
    "小小玩具": {"discount": 0.99, "invoice": 1.01},
    "家家乐":   {"discount": 1.00, "invoice": 1.01},
    "辉鸿":     {"discount": 0.98, "invoice": 1.00},
}
SEED_NO_INVOICE = ["新晨", "澄海玩具", "恒盛", "多乐", "先坤", "米可", "瑞熙", "魅力", "有利", "琳雅"]
SEED_OLD_DIR = r"C:\Users\Administrator\Desktop\按店铺对账"

STATUS_NO_INVOICE = "不开票"
STATUS_NO_PLUS    = "开票不加点"
STATUS_PLUS       = "开票加一个点"
STATUSES = [STATUS_NO_INVOICE, STATUS_NO_PLUS, STATUS_PLUS]


def config_path():
    """配置文件位置：开发时在脚本同目录，打包后在 exe 同目录"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "supplier_config.json"
    return Path(__file__).parent / "supplier_config.json"


def _status_from_remark(remark):
    if not remark:
        return None
    if "不开票" in remark:
        return STATUS_NO_INVOICE
    if "加一个点" in remark:
        return STATUS_PLUS
    if "不加点" in remark:
        return STATUS_NO_PLUS
    return STATUS_NO_PLUS


def _status_from_invoice(invoice):
    try:
        return STATUS_PLUS if float(invoice) > 1.0001 else STATUS_NO_PLUS
    except (TypeError, ValueError):
        return None


def _scan_old_summaries(old_dir):
    """扫描旧「按店铺对账」目录全部 xlsx 的「汇总」Sheet，收集每家供应商的
    收款信息 / 不开票收款信息 / 备注。返回 {供应商名: {bank, no_bank, remark}}"""
    from openpyxl import load_workbook
    suppliers = {}
    p = Path(old_dir)
    if not p.exists():
        return suppliers
    for fp in sorted(p.glob("*.xlsx")):
        try:
            wb = load_workbook(fp, data_only=True)
        except Exception:
            continue
        try:
            if "汇总" not in wb.sheetnames:
                continue
            ws = wb["汇总"]
            hdr = {}
            for j in range(1, ws.max_column + 1):
                v = ws.cell(row=1, column=j).value
                if v is not None:
                    hdr[str(v).strip()] = j
            name_col = hdr.get("厂家", 1)
            bank_col = hdr.get("收款信息")
            nobank_col = hdr.get("不开票收款信息")
            remark_col = hdr.get("备注")
            for i in range(2, ws.max_row + 1):
                name = ws.cell(row=i, column=name_col).value
                if name is None or not str(name).strip():
                    continue
                name = str(name).strip()
                rec = suppliers.setdefault(name, {})
                if bank_col:
                    b = ws.cell(row=i, column=bank_col).value
                    if b is not None and str(b).strip():
                        rec["bank"] = str(b)
                if nobank_col:
                    nb = ws.cell(row=i, column=nobank_col).value
                    if nb is not None and str(nb).strip():
                        rec["no_bank"] = str(nb)
                if remark_col:
                    rm = ws.cell(row=i, column=remark_col).value
                    if rm is not None and str(rm).strip():
                        rec["remark"] = str(rm)
        finally:
            wb.close()
    return suppliers


def init_config(old_dir=SEED_OLD_DIR):
    """扫描旧表 + 内置种子，生成初始配置 dict（不落盘）"""
    scanned = _scan_old_summaries(old_dir)

    names = set(SEED_RULES) | set(SEED_NO_INVOICE) | set(scanned)
    suppliers = {}
    for name in names:
        name = str(name).strip()
        if not name:
            continue
        rec = scanned.get(name, {})
        # 折扣：种子规则优先，否则 1.0
        seed = SEED_RULES.get(name, {})
        try:
            discount = float(seed.get("discount", 1.0))
        except (TypeError, ValueError):
            discount = 1.0
        # 开票状态：旧表备注 > 种子不开票清单 > 种子 invoice 系数 > 默认开票不加点
        status = _status_from_remark(rec.get("remark"))
        if status is None and name in SEED_NO_INVOICE:
            status = STATUS_NO_INVOICE
        if status is None:
            status = _status_from_invoice(seed.get("invoice"))
        if status is None:
            status = STATUS_NO_PLUS
        suppliers[name] = {
            "invoice_status": status,
            "discount": round(discount, 4),
            "bank": rec.get("bank", ""),
            "no_bank": rec.get("no_bank", ""),
        }
    return {"suppliers": suppliers}


_cache = {"config": None}


def load_config(force_reload=False):
    """读取配置；不存在则初始化并落盘。带进程内缓存。"""
    if _cache["config"] is not None and not force_reload:
        return _cache["config"]
    path = config_path()
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            cfg.setdefault("suppliers", {})
            _cache["config"] = cfg
            return cfg
        except (json.JSONDecodeError, OSError):
            pass
    cfg = init_config()
    save_config(cfg)
    _cache["config"] = cfg
    return cfg


def save_config(cfg):
    """落盘配置（UTF-8，无 BOM，缩进可读）"""
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    _cache["config"] = cfg


def get_supplier(name):
    """返回某供应商配置项（含默认值）；不存在返回 None"""
    cfg = load_config()
    return cfg.get("suppliers", {}).get(str(name).strip())


def get_status(name):
    """开票状态三档：不开票 / 开票不加点 / 开票加一个点"""
    s = get_supplier(name)
    if not s:
        return STATUS_NO_PLUS
    return s.get("invoice_status") or STATUS_NO_PLUS


def is_no_invoice(name):
    return get_status(name) == STATUS_NO_INVOICE


def get_rule(name):
    """金额口径：{discount, invoice}。invoice 由开票状态派生。"""
    s = get_supplier(name)
    discount = 1.0
    status = STATUS_NO_PLUS
    if s:
        try:
            discount = float(s.get("discount", 1.0))
        except (TypeError, ValueError):
            discount = 1.0
        status = s.get("invoice_status") or STATUS_NO_PLUS
    invoice = 1.01 if status == STATUS_PLUS else 1.00
    return {"discount": discount, "invoice": invoice}