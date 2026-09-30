"""打包脚本：PyArmor 加密 auth.py/main.py -> PyInstaller onefile -> 还原源码"""
import os
import sys
import glob
import shutil
import subprocess

ROOT = r"c:\Users\Administrator\Documents\order-to-excel"
BAK = os.path.join(ROOT, "_bak")
ENC = os.path.join(ROOT, "_enc")
BUILD = os.path.join(ROOT, "_build")
OUT = r"C:\Users\Administrator\Desktop\导出"
NAME = "采购单拆解生成"


def main():
    os.chdir(ROOT)

    # ---- 0. 清掉上次残留 ----
    for d in (BAK, ENC, BUILD):
        if os.path.exists(d):
            shutil.rmtree(d, ignore_errors=True)

    # ---- 1. 全量备份所有 .py ----
    os.makedirs(BAK, exist_ok=True)
    for f in glob.glob(os.path.join(ROOT, "*.py")):
        shutil.copy2(f, os.path.join(BAK, os.path.basename(f)))
    print(f"[1/4] 已备份 {len(glob.glob(os.path.join(BAK, '*.py')))} 个 .py 到 _bak/")

    try:
        # ---- 2. PyArmor 加密（输出到独立目录，不动源码）----
        os.makedirs(ENC, exist_ok=True)
        r = subprocess.run(
            [sys.executable, "-m", "pyarmor.cli", "gen", "-O", ENC, "auth.py", "main.py"],
            cwd=ROOT, capture_output=True, text=True,
        )
        print("[2/4] PyArmor rc=", r.returncode)
        print(r.stdout[-1500:] if r.stdout else "")
        if r.returncode != 0:
            print(r.stderr[-2000:])
            raise SystemExit("PyArmor 加密失败")
        if not os.path.exists(os.path.join(ENC, "main.py")):
            raise SystemExit("PyArmor 未产出 _enc/main.py")

        # ---- 3. PyInstaller ----
        os.makedirs(OUT, exist_ok=True)
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--noconfirm", "--clean",
            "--onefile", "--windowed",
            "--name", NAME,
            "--distpath", OUT,
            "--workpath", BUILD,
            "--specpath", BUILD,
            "--paths", ENC,
            "--paths", ROOT,
            # 关键：main.py 被 PyArmor 混淆后，PyInstaller 看不到里面的 import，
            # 所有业务模块必须显式声明为 hidden-import，否则不会打进去
            "--hidden-import", "auth",
            "--hidden-import", "gui",
            "--hidden-import", "excel_processor",
            "--hidden-import", "yc_processor",
            "--hidden-import", "config",
            "--hidden-import", "uploader",
            "--hidden-import", "winreg",
            "--hidden-import", "pyarmor_runtime_000000",
            # COM / 图像 / 表格
            "--hidden-import", "win32com",
            "--hidden-import", "win32com.client",
            "--hidden-import", "pythoncom",
            "--hidden-import", "openpyxl",
            "--hidden-import", "pandas",
            "--hidden-import", "xlrd",
            "--hidden-import", "PIL",
            "--collect-all", "pandas",
            "--collect-all", "openpyxl",
            os.path.join(ENC, "main.py"),
        ]
        print("[3/4] PyInstaller 开始打包（约需几分钟）...")
        r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, errors="replace")
        print("    rc =", r.returncode)
        tail = (r.stdout or "")[-2500:]
        print(tail)
        if r.returncode != 0:
            print("--- STDERR ---")
            print((r.stderr or "")[-3000:])
            raise SystemExit("PyInstaller 打包失败")

    finally:
        # ---- 4. 还原源码（无论成败）----
        restored = 0
        for f in glob.glob(os.path.join(BAK, "*.py")):
            dst = os.path.join(ROOT, os.path.basename(f))
            shutil.copy2(f, dst)
            restored += 1
        print(f"[4/4] 已从 _bak/ 还原 {restored} 个 .py")

    exe = os.path.join(OUT, NAME + ".exe")
    if os.path.exists(exe):
        print(f"\n✅ 打包完成: {exe}  {os.path.getsize(exe) / 1024 / 1024:.1f} MB")
    else:
        print("\n❌ 未找到产物 exe")


if __name__ == "__main__":
    main()