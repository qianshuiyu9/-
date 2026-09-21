import subprocess, os, sys
sys.stdout.reconfigure(encoding='utf-8')

exe = os.path.join(os.path.expanduser("~"), "Desktop", "导出", "采购单拆解生成.exe")
today_code = "455615"

print("测试: 正确验证码通过（加密版）")
try:
    proc = subprocess.run([exe, "--no-upload"], input=today_code + "\n", capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=90)
    out = proc.stdout + proc.stderr
    for line in out.splitlines()[:50]:
        print(f"  {line}")
    print(f"\n  exit={proc.returncode}")
except subprocess.TimeoutExpired:
    print("  TIMEOUT")
