import re
import os
from playwright.sync_api import sync_playwright
from config import PLATFORM_URL, PLATFORM_USERNAME, PLATFORM_PASSWORD, UPLOAD_CONFIG


def extract_supplier_name(filename):
    name = os.path.splitext(filename)[0]
    name = re.sub(r"[\-_].*$", "", name)
    name = re.sub(r"旗舰店|专卖店|专营店|超市|店铺", "", name)
    return name.strip()


class PlatformUploader:
    def __init__(self, headless=False):
        self.headless = headless
        self.playwright = None
        self.browser = None
        self.page = None
        self.c = UPLOAD_CONFIG

    def start(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=self.headless)
        self.page = self.browser.new_page()

    def login(self):
        c = self.c
        print(f"  打开登录页: {PLATFORM_URL}")
        self.page.goto(PLATFORM_URL)
        self.page.wait_for_load_state("networkidle")

        if PLATFORM_USERNAME:
            try:
                self.page.fill(c["login_user"], PLATFORM_USERNAME)
            except Exception:
                self.page.get_by_role("textbox").first.fill(PLATFORM_USERNAME)

        if PLATFORM_PASSWORD:
            try:
                self.page.fill(c["login_pass"], PLATFORM_PASSWORD)
            except Exception:
                pwd_inputs = self.page.locator("input[type='password']")
                pwd_inputs.first.fill(PLATFORM_PASSWORD)

        try:
            self.page.click(c["login_btn"])
        except Exception:
            self.page.get_by_role("button").first.click()

        self.page.wait_for_load_state("networkidle")
        print("  ✓ 登录完成")

    def go_to_purchase_import(self):
        c = self.c
        print("  进入货品采购 → 导入采购表格")

        try:
            self.page.click(c["menu_goods"])
            self.page.wait_for_timeout(300)
        except Exception:
            pass

        try:
            self.page.click(c["menu_purchase"])
            self.page.wait_for_load_state("networkidle")
        except Exception as e:
            print(f"    [注意] 菜单点击: {e}")

        try:
            self.page.click(c["btn_import"], timeout=5000)
        except Exception:
            try:
                self.page.click(c["btn_new_plan"], timeout=3000)
                self.page.wait_for_timeout(500)
                self.page.click(c["btn_import_table_tab"], timeout=3000)
            except Exception as e:
                print(f"    [注意] 导入按钮: {e}")

        self.page.wait_for_load_state("networkidle")

    def select_supplier(self, supplier_name):
        c = self.c
        print(f"  选择供应商: {supplier_name}")

        try:
            supplier_select = self.page.locator(".el-select").first
            supplier_select.click()
            self.page.wait_for_timeout(300)

            search_input = self.page.locator(".el-select-dropdown input").first
            search_input.fill(supplier_name)
            self.page.wait_for_timeout(500)

            option = self.page.get_by_text(supplier_name, exact=True).first
            option.click()
            self.page.wait_for_timeout(300)
            print(f"    ✓ 已选择: {supplier_name}")
        except Exception as e:
            print(f"    [警告] 供应商选择失败，跳过: {e}")

    def upload_single(self, file_path, supplier_name=""):
        c = self.c
        display_name = supplier_name or os.path.basename(file_path)
        print(f"\n  📤 上传: {display_name}")

        self.go_to_purchase_import()

        if supplier_name:
            self.select_supplier(supplier_name)

        file_input = self.page.locator(c["file_upload"]).first
        file_input.set_input_files(file_path)
        print(f"    ✓ 文件已选择")

        try:
            self.page.click(c["file_submit"], timeout=10000)
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(1000)
            print(f"    ✓ 上传完成")
        except Exception as e:
            print(f"    [警告] 提交按钮点击失败: {e}")

    def upload_all(self, file_results):
        print("\n" + "=" * 60)
        print("步骤 2: 自动登录并上传到集成平台")
        print("=" * 60)

        self.start()
        try:
            self.login()
            for item in file_results:
                supplier = item.get("supplier") or extract_supplier_name(item["name"])
                try:
                    self.upload_single(item["output"], supplier)
                except Exception as e:
                    print(f"    ✗ 上传失败: {e}")
        finally:
            self.close()

    def close(self):
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        print("  浏览器已关闭")


if __name__ == "__main__":
    uploader = PlatformUploader(headless=False)
    uploader.start()
    uploader.login()
    uploader.close()
