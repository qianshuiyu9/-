import sys
import argparse
# Windows 默认 GBK 控制台无法显示 emoji，强制切 UTF-8
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from excel_processor import process_all
from yc_processor import process_yuanchuang_all
from uploader import PlatformUploader


def run(upload=True, headless=False):
    print("\n🚀 订单自动化处理流程启动\n")

    file_results = process_all()

    # 原创厂家独立管线（放最后，避免被 process_all 的目录清理擦掉）
    yc_results = process_yuanchuang_all()

    # 合并结果用于上传
    all_results = file_results or []
    if yc_results:
        all_results.extend(yc_results)

    if not all_results:
        print("\n[退出] 没有可处理的文件")
        return

    if upload:
        uploader = PlatformUploader(headless=headless)
        uploader.upload_all(all_results)
    else:
        print("\n[跳过上传] 仅生成本地表格")

    print("\n✅ 全部完成！\n")


def main():
    parser = argparse.ArgumentParser(description="订单表格处理 → 自动上传")
    parser.add_argument("--no-upload", action="store_true", help="仅生成本地表格，不上传")
    parser.add_argument("--headless", action="store_true", help="浏览器无界面模式")
    args = parser.parse_args()

    run(upload=not args.no_upload, headless=args.headless)


if __name__ == "__main__":
    main()
