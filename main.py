import sys
import argparse
import os
from datetime import datetime

# Windows 默认 GBK 控制台无法显示 emoji，强制切 UTF-8
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# ⚠️ 先只 import 轻量的 auth（不要碰 numpy/pandas）
# 在验证通过后才 import excel_processor 等重模块
from auth import prompt_and_verify, gen_code


# ---------- 下面两个函数需要 OUTPUT_DIR，延迟 import ----------
def print_summary(all_results, output_dir=None):
    from datetime import datetime as _dt
    """处理完成后打印汇总日志，并写入桌面日志文件。"""
    output_dir = output_dir or ""
    print("\n" + "=" * 70)
    print("处理汇总日志")
    print("=" * 70)
    print(f"时间: {_dt.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"输出目录: {output_dir}")
    print("-" * 70)
    print(f"{'厂家':<10} {'文件名':<45} {'行':>4} {'规格拆解':>6} {'图片':>4}")
    print("-" * 70)

    # 按厂家分组汇总
    supplier_totals = {}
    for r in all_results:
        supplier = r.get("supplier", "未知")
        name = r.get("name", "")
        rows = r.get("rows", 0)
        specs = r.get("spec_changes", 0)
        imgs = r.get("image_count", 0)
        print(f"  {supplier:<8} {name[:45]:<45} {rows:>4} {specs:>6} {imgs:>4}")
        if supplier not in supplier_totals:
            supplier_totals[supplier] = {"files": 0, "rows": 0, "specs": 0, "imgs": 0}
        supplier_totals[supplier]["files"] += 1
        supplier_totals[supplier]["rows"] += rows
        supplier_totals[supplier]["specs"] += specs
        supplier_totals[supplier]["imgs"] += imgs

    print("-" * 70)
    total_files = sum(v["files"] for v in supplier_totals.values())
    total_rows = sum(v["rows"] for v in supplier_totals.values())
    total_specs = sum(v["specs"] for v in supplier_totals.values())
    total_imgs = sum(v["imgs"] for v in supplier_totals.values())
    print(f"  {'合计':<8} {'':<45} {total_rows:>4} {total_specs:>6} {total_imgs:>4}")

    # 按厂家分组打印子汇总
    print()
    for supplier, t in supplier_totals.items():
        print(f"  [{supplier}] {t['files']}文件 / {t['rows']}行 / 规格拆解{t['specs']} / 图片{t['imgs']}")

    print("=" * 70)

    # 写入日志文件（桌面导出目录）
    try:
        log_dir = os.path.join(os.path.expanduser("~"), "Desktop", "导出")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, f"处理日志_{_dt.now().strftime('%Y%m%d_%H%M%S')}.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"处理时间: {_dt.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"输出目录: {output_dir}\n\n")
            f.write(f"{'厂家':<10} {'文件名':<50} {'行':>4} {'规格拆解':>6} {'图片':>4}\n")
            f.write("-" * 80 + "\n")
            for r in all_results:
                supplier = r.get("supplier", "未知")
                name = r.get("name", "")
                rows = r.get("rows", 0)
                specs = r.get("spec_changes", 0)
                imgs = r.get("image_count", 0)
                f.write(f"{supplier:<10} {name[:50]:<50} {rows:>4} {specs:>6} {imgs:>4}\n")
            f.write("-" * 80 + "\n")
            f.write(f"合计: {total_files}文件 / {total_rows}行 / 规格拆解{total_specs} / 图片{total_imgs}\n")
            for supplier, t in supplier_totals.items():
                f.write(f"  [{supplier}] {t['files']}文件 / {t['rows']}行 / 规格拆解{t['specs']} / 图片{t['imgs']}\n")
        print(f"日志已保存: {log_path}")
    except Exception as e:
        print(f"[警告] 日志文件写入失败: {e}")


def run(upload=True, headless=False):
    # 先过验证码 —— 通过后才 import 重模块
    if not prompt_and_verify():
        sys.exit(1)

    # ⬇️ 延迟 import 重模块（numpy/pandas/openpyxl）
    from excel_processor import process_all, OUTPUT_DIR
    from yc_processor import process_yuanchuang_all
    from uploader import PlatformUploader

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

    print_summary(all_results, output_dir=OUTPUT_DIR)
    print("\n✅ 全部完成！\n")


def main():
    parser = argparse.ArgumentParser(description="订单表格处理 → 自动上传")
    parser.add_argument("--no-upload", action="store_true", help="仅生成本地表格，不上传")
    parser.add_argument("--headless", action="store_true", help="浏览器无界面模式")
    parser.add_argument("--gen-code", action="store_true", help="[发布者工具] 打印今日验证码后退出")
    args = parser.parse_args()

    if args.gen_code:
        from datetime import date
        print("=" * 40)
        print("  验证码计算器（发布者工具）")
        print("=" * 40)
        today = date.today()
        for offset in range(-2, 3):
            d = date.fromordinal(today.toordinal() + offset)
            code = gen_code(d)
            mark = "  ← 今天" if offset == 0 else ""
            print(f"  {d.strftime('%Y-%m-%d')}: {code:06d}{mark}")
        print("=" * 40)
        return

    run(upload=False, headless=args.headless)


if __name__ == "__main__":
    main()
    try:
        input("\n按回车键退出...")
    except (EOFError, KeyboardInterrupt):
        pass
