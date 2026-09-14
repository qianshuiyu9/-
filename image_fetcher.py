"""供应商图片抓取模块。

当源表格没有货品图片时（如"原创""家家乐"等厂家），
自动打开其对应的微信小程序/H5 页面，按商品名搜索并下载图片。

支持两种抓取模式（由 config.SUPPLIER_IMAGE_SOURCES[*].search_mode 决定）：
  - "url" / "form"：通过 Playwright 打开 H5 网页搜索（需配置 search_url）
  - "wechat"：通过 uiautomation 操作微信 PC 版打开小程序搜索（需配置 miniprogram_name）
"""
import re
import io
import glob
import os
import urllib.parse
from playwright.sync_api import sync_playwright
from wechat_fetcher import WeChatMiniProgramFetcher


def _find_chrome_exe():
    """在 Playwright 浏览器目录中查找 chrome.exe（优先用完整 Chromium，避免依赖 headless shell）。"""
    base = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ms-playwright")
    matches = glob.glob(os.path.join(base, "chromium-*", "chrome-win*", "chrome.exe"))
    if matches:
        return matches[0]
    return None


def _resolve_image_url(src, page_url):
    """把图片地址补全为绝对 URL。"""
    if not src:
        return None
    src = src.strip()
    if src.startswith("http://") or src.startswith("https://"):
        return src
    if src.startswith("//"):
        return "https:" + src
    # 相对路径，基于当前页面 URL 拼接
    base = urllib.parse.urljoin(page_url, src)
    return base


class MiniProgramImageFetcher:
    """供应商图片抓取器。

    根据 cfg["search_mode"] 自动选择后端：
      - "wechat"：WeChatMiniProgramFetcher（操作微信 PC 版小程序）
      - 其他：Playwright 打开 H5 页面
    """

    def __init__(self, headless=True):
        self.headless = headless
        self._wechat_fetcher = None
        # Playwright 相关
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    def _use_wechat(self, cfg):
        return cfg and cfg.get("search_mode") == "wechat"

    def start(self, cfg=None):
        if self._use_wechat(cfg):
            # 微信模式：无需启动 Playwright
            return
        self.playwright = sync_playwright().start()
        chrome_exe = _find_chrome_exe()
        launch_kwargs = {"headless": self.headless}
        if chrome_exe:
            # 直接用完整 Chromium，避免依赖单独的 headless shell
            launch_kwargs["executable_path"] = chrome_exe
        self.browser = self.playwright.chromium.launch(**launch_kwargs)
        # 模拟手机端，适配 H5/小程序页面
        self.context = self.browser.new_context(
            user_agent=(
                "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) "
                "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 "
                "Mobile/15E148 Safari/604.1"
            ),
            viewport={"width": 390, "height": 844},
        )
        self.page = self.context.new_page()

    def _extract_image_url(self, cfg):
        """从当前页面提取第一个商品的图片 URL。"""
        sel = cfg["selectors"]
        try:
            # 定位到第一个商品容器
            product = self.page.locator(sel["product_list"]).first
            product.wait_for(state="visible", timeout=cfg.get("timeout", 10000))
            img = product.locator(sel["product_image"]).first
            # 优先取 data-src（懒加载），其次 src
            src = img.get_attribute("data-src") or img.get_attribute("src")
            if src:
                return _resolve_image_url(src, self.page.url)
        except Exception:
            pass
        return None

    def fetch_image(self, keyword, cfg):
        """搜索商品并返回图片二进制数据；失败返回 None。

        Args:
            keyword: 商品名称
            cfg: SUPPLIER_IMAGE_SOURCES 中的供应商配置
        """
        if not cfg or not cfg.get("enabled"):
            return None

        # 微信小程序模式
        if self._use_wechat(cfg):
            mp_name = cfg.get("miniprogram_name")
            if not mp_name:
                return None
            if self._wechat_fetcher is None:
                self._wechat_fetcher = WeChatMiniProgramFetcher(
                    miniprogram_name=mp_name,
                    search_box_pos=cfg.get("search_box_pos", (200, 50)),
                    crop_region=cfg.get("crop_region", (10, 110, 190, 290)),
                    timeout=cfg.get("timeout", 15000) // 1000,
                )
            try:
                return self._wechat_fetcher.fetch_image(keyword)
            except Exception as e:
                print(f"    [微信抓取] 异常: {e}")
                return None

        # H5 模式
        if not cfg.get("search_url"):
            return None
        if "{keyword}" not in cfg["search_url"] and cfg.get("search_mode") != "form":
            return None

        sel = cfg["selectors"]
        encoded = urllib.parse.quote(keyword)

        try:
            if cfg.get("search_mode") == "form":
                self.page.goto(cfg["search_url"], wait_until="domcontentloaded")
                self.page.wait_for_load_state("networkidle")
                self.page.fill(sel["search_input"], keyword)
                self.page.click(sel["search_button"])
            else:
                url = cfg["search_url"].replace("{keyword}", encoded)
                self.page.goto(url, wait_until="domcontentloaded")

            self.page.wait_for_load_state("networkidle")
            # 稍等结果渲染
            self.page.wait_for_timeout(800)

            img_url = self._extract_image_url(cfg)
            if not img_url:
                return None

            # 下载图片
            resp = self.page.request.get(img_url)
            if resp.ok:
                return resp.body()
        except Exception as e:
            print(f"    [图片抓取] 异常: {e}")
        return None

    def fetch_images(self, keywords, cfg, on_progress=None):
        """批量搜索多个商品的图片。

        Args:
            keywords: 商品名称列表
            cfg: 供应商配置
            on_progress: 回调函数 (index, keyword, ok) -> None
        Returns:
            dict: {商品名: 图片bytes 或 None}
        """
        results = {}
        for i, kw in enumerate(keywords):
            data = self.fetch_image(kw, cfg) if kw else None
            results[kw] = data
            if on_progress:
                on_progress(i, kw, data is not None)
        return results

    def close(self):
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        self._wechat_fetcher = None


def bytes_to_openpyxl_image(data, size=(80, 80)):
    """把图片二进制转成 openpyxl Image 对象（用 PIL 缩放字节，确保保存后尺寸持久）。"""
    from openpyxl.drawing.image import Image
    try:
        from excel_processor import _resize_image_bytes
        data = _resize_image_bytes(data, size[0], size[1])
        img = Image(io.BytesIO(data))
        return img
    except Exception:
        return None
