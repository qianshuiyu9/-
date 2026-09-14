PLATFORM_URL = "https://live.catchtoy.cn/qterp/mportal/index.html#/login"
PLATFORM_USERNAME = "18728023881"
PLATFORM_PASSWORD = "654321"

# 输入目录：桌面/原始数据文件/{厂家名}/{厂家给的原始表格}
# 输出目录：桌面/处理后标准表格/{厂家名}/{厂家名+店铺或收货人}.xlsx
import os as _os
_DESKTOP = _os.path.join(_os.path.expanduser("~"), "Desktop")
INPUT_DIR = _os.path.join(_DESKTOP, "原始数据文件")
OUTPUT_DIR = _os.path.join(_DESKTOP, "处理后标准表格")

TARGET_COLUMNS = [
    "类别",
    "名称",
    "图片",
    "单位",
    "进货数量",
    "单价",
    "条码",
    "规格",
    "用途",
    "标签",
    "积分倍数",
    "供应商商品编码",
    "尺寸",
    "展示名",
    "总金额",
]

REQUIRED_COLUMNS = [
    "名称", "单位", "进货数量", "单价",
    "规格", "用途", "标签", "积分倍数",
]

COLUMN_MAPPING = {
    "类别":       ["类别", "分类", "商品分类", "品类"],
    "名称":       ["商品名称", "名称", "名字", "货品", "货品名称", "商品", "品名", "产品", "产品名称", "商品全名"],
    "单位":       ["单位", "计量单位", "Unit", "unit"],
    "进货数量":   ["进货数量", "数量", "件数", "订购数量", "订货数量", "总数量", "Qty", "qty"],
    "单价":       ["单价", "含税单价", "折后价", "价格", "售价", "批发价", "进货价", "Price", "price"],
    "条码":       ["条码", "条形码", "商品条码", "Barcode", "barcode", "EAN", "ean"],
    "规格":       ["规格", "型号", "产品规格", "商品规格"],
    "用途":       ["用途", "使用说明"],
    "标签":       ["标签", "Tag", "tag"],
    "供应商商品编码": ["供应商商品编码", "商品编码", "SKU", "sku", "编码"],
    "尺寸":       ["尺寸", "大小", "Size", "size"],
    "总金额":     ["合计金额", "含税金额", "金额", "总金额", "小计", "总价"],
}

# 数量单位（匹配到会参与数量换算：数量 × 倍数，单价 ÷ 倍数）
SPEC_UNITS = [
    "入", "支", "个", "瓶", "盒", "箱", "包", "袋", "只", "片", "件",
    "罐", "听", "条", "卷", "本", "张", "块", "粒", "颗", "枚",
    "打", "套", "组", "对", "双", "副",
]

# 体积/重量单位（仅写入规格文本，不参与数量换算）
VOLUME_WEIGHT_UNITS = [
    "ml", "毫升", "l", "升", "kg", "千克", "g", "克", "斤", "公斤", "吨",
]

# 按长度降序排列，保证长单位（如"罐装"）优先于短单位匹配
_SPEC_UNITS_SORTED = sorted(SPEC_UNITS, key=len, reverse=True)
_SPEC_UNITS_ALT = "|".join(_SPEC_UNITS_SORTED)
_VW_UNITS_ALT = "|".join(VOLUME_WEIGHT_UNITS)

# 特定单位的固定换算值（如"1打"=12个，"1对"=2个）
# 实际倍数 = 数字 × 单位固定值
UNIT_FIXED_VALUE = {
    "打": 12,
    "对": 2,
    "双": 2,
    "副": 2,
}

# 乘法表达式：如 1x24 / 2*6 / 1×12 / 1x24包（尾部可选单位）
_SPEC_MULT = r"(?P<mult>(?<![0-9])(?P<mult_a>\d+(?:\.\d+)?)\s*[xX\*×]\s*(?P<mult_b>\d+(?:\.\d+)?)(?:\s*(?P<mult_unit>" + _SPEC_UNITS_ALT + r")[装裝]?)?)"
# 加法表达式（赠品）：如 6+1入 / 10+2支，倍数取第一个数（+后为赠品不计入）
_SPEC_ADD = r"(?P<add>(?<![0-9])(?P<add_a>\d+(?:\.\d+)?)\s*\+\s*\d+(?:\.\d+)?(?:\s*(?P<add_unit>" + _SPEC_UNITS_ALT + r")[装裝]?)?)"
# 数量+单位（可选"装/裝"后缀）：如 12支 / 6入 / 24瓶装 / 1罐装
# (?<![0-9\-]) 确保数字前不是数字或减号，避免从编码(如007-295卷)中误匹配出295卷
_SPEC_COUNT = r"(?P<count>(?<![0-9\-])(?P<count_num>\d+(?:\.\d+)?)\s*(?P<count_unit>" + _SPEC_UNITS_ALT + r")[装裝]?)"
# 体积/重量：如 500ml / 1kg / 330毫升
_SPEC_VW = r"(?P<vw>(?<![0-9])(?P<vw_num>\d+(?:\.\d+)?)\s*(?P<vw_unit>" + _VW_UNITS_ALT + r"))"

# 复合正则：乘法 > 加法(赠品) > 数量单位 > 体积重量，非重叠匹配
SPEC_REGEX = _SPEC_MULT + "|" + _SPEC_ADD + "|" + _SPEC_COUNT + "|" + _SPEC_VW

SKIP_SPEC_KEYWORDS = ["拆散", "散装", "单支", "单个", "单件", "单品", "洞洞乐"]

FIXED_INTEGRAL_MULTIPLIER = 3.5

# 输出表格的默认字段值（源表格未提供时使用）
# 规格：商品类别；用途/标签/单位：平台默认分类
DEFAULT_FIELD_VALUES = {
    "规格": "玩具类",
    "用途": "兑换, 零售",
    "标签": "儿童玩具",
    "单位": "个",
}

UPLOAD_CONFIG = {
    "login_user": "#username",
    "login_pass": "#password",
    "login_btn": "button[type='submit']",

    "menu_goods": "text=货品",
    "menu_purchase": "text=货品采购",

    "btn_import": "button:has-text('导入采购表格')",

    "supplier_selector": "input[type='text']",
    "supplier_option": ".el-select-dropdown__item",

    "file_upload": "input[type='file']",
    "file_submit": "button:has-text('确定')",

    "btn_new_plan": "button:has-text('新建采购计划')",
    "btn_import_table_tab": "text=导入采购表格",
}

# ---------------------------------------------------------------------------
# 供应商图片抓取配置
# 当厂家（由文件名识别）出现在此配置中时，若源表格没有货品图片，
# 会自动打开其对应的微信小程序搜索商品并截图填入标准表格。
#
# search_mode 说明：
#   "wechat" —— 通过 uiautomation 操作微信 PC 版打开小程序搜索（推荐，无需 URL）
#              需要：本机已安装并登录微信 PC 版
#   "url"    —— Playwright 打开 H5 搜索页，URL 中含 {keyword} 占位符
#   "form"   —— Playwright 打开搜索页，填搜索框再点按钮
#
# 微信模式下的关键参数：
#   miniprogram_name  : 小程序在微信中的名称（如"原创优品生活馆"）
#   search_box_pos    : 小程序内搜索框的相对坐标 (x, y)，默认 (200, 50)
#   crop_region       : 搜索结果首张商品图的截取区域 (左, 上, 右, 下)，默认 (10, 110, 190, 290)
#                       （可根据实际小程序布局调整）
#   若搜索不到商品，自动跳过该款货品（不报错、不中断）
# ---------------------------------------------------------------------------
SUPPLIER_IMAGE_SOURCES = {
    "原创": {
        "enabled": False,
        "search_mode": "wechat",
        "miniprogram_name": "原创优品生活馆",
        "search_box_pos": (200, 50),
        "crop_region": (10, 110, 190, 290),
        "image_size": (80, 80),
        "timeout": 15000,
    },
    "家家乐": {
        "enabled": False,
        "search_mode": "wechat",
        "miniprogram_name": "家家乐礼品汇",
        "search_box_pos": (200, 50),
        "crop_region": (10, 110, 190, 290),
        "image_size": (80, 80),
        "timeout": 15000,
    },
}
