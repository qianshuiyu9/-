"""微信 PC 版小程序图片抓取。

通过 uiautomation 操作微信 PC 客户端：
  1. 在微信中搜索并打开指定小程序
  2. 在小程序内搜索商品
  3. 截图并截取首张商品图片
若任何步骤失败或搜索无结果，返回 None（跳过该商品）。

注意：需要用户本机已安装并登录微信 PC 版。
"""
import io
import time
import os
import subprocess

try:
    import uiautomation as auto
    HAS_UIA = True
except ImportError:
    HAS_UIA = False

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


WECHAT_PATHS = [
    r"C:\Program Files\Tencent\WeChat\WeChat.exe",
    r"C:\Program Files (x86)\Tencent\WeChat\WeChat.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Tencent", "WeChat", "WeChat.exe"),
]


def _find_wechat_exe():
    for p in WECHAT_PATHS:
        if p and os.path.exists(p):
            return p
    return None


def _find_wechat_window(timeout=5):
    """查找微信主窗口。"""
    # 微信主窗口标题通常是"微信"
    win = auto.WindowControl(searchDepth=1, Name="微信")
    if win.Exists(timeout):
        return win
    # 尝试按进程名查找
    win = auto.WindowControl(searchDepth=1, ClassName="WeChatMainWndForPC")
    if win.Exists(1):
        return win
    return None


def _launch_wechat():
    """尝试启动微信。"""
    exe = _find_wechat_exe()
    if exe:
        subprocess.Popen([exe])
        return True
    return False


class WeChatMiniProgramFetcher:
    """通过微信 PC 版打开小程序搜索商品图片。"""

    def __init__(self, miniprogram_name, search_box_pos=(200, 50),
                 crop_region=(10, 110, 190, 290), timeout=15):
        """
        Args:
            miniprogram_name: 小程序名称（如"原创优品生活馆"）
            search_box_pos: 小程序内搜索框的相对坐标 (x, y)
            crop_region: 搜索结果中首张商品图的截取区域 (left, top, right, bottom)
            timeout: 各步骤超时（秒）
        """
        self.miniprogram_name = miniprogram_name
        self.search_box_pos = search_box_pos
        self.crop_region = crop_region
        self.timeout = timeout
        self._mp_window = None

    def _activate_wechat(self):
        """激活微信主窗口，返回窗口对象或 None。"""
        win = _find_wechat_window(timeout=3)
        if not win:
            if _launch_wechat():
                win = _find_wechat_window(timeout=15)
        if not win:
            return None
        try:
            win.SetActive()
            win.SetTopmost(True)
            time.sleep(0.5)
            win.SetTopmost(False)
            return win
        except Exception:
            return None

    def _open_miniprogram(self, wechat_win):
        """在微信中搜索并打开小程序，返回小程序窗口或 None。"""
        try:
            # Ctrl+F 聚焦搜索框（微信搜索全局快捷键）
            wechat_win.SendKeys("{Ctrl}f")
            time.sleep(0.6)
            # 清空后输入小程序名
            wechat_win.SendKeys("{Ctrl}a{Del}")
            wechat_win.SendKeys(self.miniprogram_name, waitTime=0.05)
            time.sleep(1.5)

            # 尝试在搜索结果中找到小程序项并点击
            clicked = self._click_mp_in_results(wechat_win)
            if not clicked:
                # 兜底1：向下选中第一个结果再回车
                wechat_win.SendKeys("{Down}")
                time.sleep(0.3)
                wechat_win.SendKeys("{Enter}")

            # 等待小程序窗口出现（轮询最多 8 秒）
            mp_win = None
            for _ in range(16):
                mp_win = self._find_mp_window()
                if mp_win:
                    break
                time.sleep(0.5)
            return mp_win
        except Exception:
            return None

    def _click_mp_in_results(self, wechat_win):
        """在搜索结果下拉中查找并点击小程序项。"""
        name = self.miniprogram_name
        # 尝试精确匹配文本
        for depth in range(1, 6):
            ctrl = wechat_win.TextControl(searchDepth=depth, Name=name)
            if ctrl.Exists(0.5):
                ctrl.Click(simulateMove=False)
                return True
        # 尝试包含匹配
        for depth in range(1, 8):
            try:
                items = wechat_win.GetChildren()
                # 遍历子控件找包含小程序名的文本
                found = self._find_text_contains(wechat_win, name, max_depth=8)
                if found:
                    found.Click(simulateMove=False)
                    return True
            except Exception:
                pass
            break
        return False

    def _find_text_contains(self, root, text, max_depth=8):
        """在 UI 树中查找包含指定文本的控件。"""
        try:
            for depth in range(1, max_depth + 1):
                ctrls = root.GetChildren()
                # 用 Name 包含匹配
                found = auto.TextControl(searchDepth=depth, Name=lambda s: text in (s or ""))
                if found.Exists(0.3):
                    return found
        except Exception:
            pass
        return None

    def _find_mp_window(self):
        """查找已打开的小程序窗口。"""
        name = self.miniprogram_name
        # 1. 精确匹配窗口标题
        win = auto.WindowControl(searchDepth=1, Name=name)
        if win.Exists(1):
            return win
        # 2. 包含小程序名的窗口
        win = auto.WindowControl(searchDepth=1, Name=lambda s: name in (s or ""))
        if win.Exists(1):
            return win
        # 3. 遍历所有顶层窗口找包含"小程序"或小程序名的
        try:
            root = auto.GetRootControl()
            for w in root.GetChildren():
                wname = w.Name or ""
                if name in wname or "小程序" in wname:
                    return w
        except Exception:
            pass
        return None

    def _search_product(self, mp_win, product_name):
        """在小程序窗口中搜索商品。"""
        try:
            mp_win.SetActive()
            time.sleep(0.4)
            # 优先尝试找到输入框控件
            edit = mp_win.EditControl(searchDepth=10)
            if edit.Exists(1):
                edit.Click(simulateMove=False)
                time.sleep(0.3)
                edit.SendKeys("{Ctrl}a{Del}")
                edit.SendKeys(product_name, waitTime=0.05)
                time.sleep(0.3)
                edit.SendKeys("{Enter}")
            else:
                # 兜底：点击配置的搜索框坐标
                pos = self.search_box_pos
                rect = mp_win.BoundingRectangle
                x = rect.left + pos[0]
                y = rect.top + pos[1]
                auto.Click(x, y)
                time.sleep(0.3)
                auto.SendKeys("{Ctrl}a{Del}")
                auto.SendKeys(product_name, waitTime=0.05)
                time.sleep(0.3)
                auto.SendKeys("{Enter}")
            time.sleep(1.5)
            return True
        except Exception:
            return False

    def _capture_product_image(self, mp_win):
        """截取小程序窗口中首张商品图片区域。"""
        if not HAS_PIL:
            return None
        try:
            bmp = mp_win.CaptureToImage()
            if bmp is None:
                return None
            left, top, right, bottom = self.crop_region
            # 边界保护
            w, h = bmp.size
            left = max(0, min(left, w - 1))
            top = max(0, min(top, h - 1))
            right = max(left + 1, min(right, w))
            bottom = max(top + 1, min(bottom, h))
            cropped = bmp.crop((left, top, right, bottom))
            buf = io.BytesIO()
            cropped.save(buf, format="PNG")
            return buf.getvalue()
        except Exception:
            return None

    def fetch_image(self, product_name):
        """搜索商品并返回图片 bytes；搜索不到或失败返回 None。"""
        if not HAS_UIA:
            return None

        wechat = self._activate_wechat()
        if not wechat:
            return None

        mp_win = self._open_miniprogram(wechat)
        if not mp_win:
            return None

        if not self._search_product(mp_win, product_name):
            return None

        # 等待搜索结果渲染
        time.sleep(1.5)

        data = self._capture_product_image(mp_win)
        if data and len(data) > 200:
            return data
        return None
