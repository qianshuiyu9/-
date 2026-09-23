"""
采购单拆解生成 —— tkinter 图形界面

特性：
- 从 INPUT_DIR 子文件夹动态扫描厂家列表
- 厂家多选（勾选），全选 / 清空
- 开始处理在子线程运行，主线程用 Queue+after 刷新日志
- 日志同时输出到 GUI Text 和 sys.stdout
- 处理完成后可继续选厂家跑下一轮
"""
import os
import sys
import queue
import threading
import subprocess

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ---------- 配置 ----------
APP_TITLE = "采购单拆解生成 v2.0"
APP_WIDTH = 720
APP_HEIGHT = 560


# =============================================================
# 日志重定向：让 print() 同时输出到 GUI 和控制台
# =============================================================
class _GuiStdout:
    def __init__(self, q):
        self._q = q
        self._console = sys.__stdout__  # --windowed 下可能为 None
        self._buf = ""

    def write(self, s):
        if not s:
            return
        if self._console is not None:
            try:
                self._console.write(s)
            except Exception:
                pass
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            self._q.put(line)

    def flush(self):
        if self._console is not None:
            try:
                self._console.flush()
            except Exception:
                pass


class _GuiStderr:
    def __init__(self, q):
        self._q = q
        self._console = sys.__stderr__  # --windowed 下可能为 None

    def write(self, s):
        if not s:
            return
        if self._console is not None:
            try:
                self._console.write(s)
            except Exception:
                pass
        for line in s.splitlines():
            self._q.put(line)

    def flush(self):
        if self._console is not None:
            try:
                self._console.flush()
            except Exception:
                pass


# =============================================================
# 主窗口
# =============================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
        self.root.minsize(640, 480)

        # 验证码是否已通过（启动时先让用户输码）
        self.auth_ok = False

        # 厂家列表（从 INPUT_DIR 动态扫描）
        self._suppliers = []
        self._check_vars = {}  # {厂家名: BooleanVar}

        # 处理线程锁
        self._running = False

        # 日志队列
        self._log_q = queue.Queue()

        # 构建界面
        self._build_ui()

        # 刷新厂家列表
        self._refresh_suppliers()

        # 启动日志轮询（每 80ms）
        self.root.after(80, self._drain_log_queue)

    # ---------------- 构建 UI ----------------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # 顶部：验证码状态
        self.auth_frame = ttk.Frame(self.root)
        self.auth_frame.pack(fill="x", **pad)
        self.auth_label = ttk.Label(self.auth_frame, text="🔐 未验证（请先点开始处理，会要求输验证码）", foreground="#888")
        self.auth_label.pack(side="left")

        # 主区域：左边厂家列表，右边按钮
        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=False, **pad)

        # 左：厂家选择
        left = ttk.LabelFrame(main, text="厂家选择（可多选）")
        left.pack(side="left", fill="both", expand=False)

        self.check_frame = ttk.Frame(left)
        self.check_frame.pack(fill="both", expand=True, padx=6, pady=6)

        # 全选 / 清空
        btn_frame = ttk.Frame(left)
        btn_frame.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Button(btn_frame, text="全选", width=8, command=self._select_all).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="清空", width=8, command=self._clear_all).pack(side="left", padx=2)

        # 刷新厂家
        ttk.Button(left, text="🔄 刷新厂家列表", command=self._refresh_suppliers).pack(
            fill="x", padx=6, pady=(0, 6)
        )

        # 右：操作按钮
        right = ttk.Frame(main)
        right.pack(side="left", fill="y", padx=(12, 0))

        self.start_btn = ttk.Button(right, text="▶ 开始处理", width=18, command=self._on_start)
        self.start_btn.pack(pady=4)

        self.open_input_btn = ttk.Button(right, text="📂 打开输入目录", width=18, command=self._open_input_dir)
        self.open_input_btn.pack(pady=4)

        self.open_output_btn = ttk.Button(right, text="📂 打开输出目录", width=18, command=self._open_output_dir)
        self.open_output_btn.pack(pady=4)

        self.clear_log_btn = ttk.Button(right, text="🗑 清空日志", width=18, command=self._clear_log)
        self.clear_log_btn.pack(pady=4)

        # 底部：日志框
        log_frame = ttk.LabelFrame(self.root, text="运行日志")
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_text = tk.Text(log_frame, wrap="word", height=14, font=("Consolas", 10))
        self.log_text.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        # 只读
        self.log_text.configure(state="disabled")

        sb = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=sb.set)

    # ---------------- 厂家列表 ----------------
    def _refresh_suppliers(self):
        """从 INPUT_DIR 扫描厂家子文件夹，重建勾选框。"""
        try:
            from config import INPUT_DIR
        except Exception:
            messagebox.showerror("错误", "无法读取 config.py 中的 INPUT_DIR")
            return

        if not os.path.isdir(INPUT_DIR):
            messagebox.showwarning("提示", f"输入目录不存在：\n{INPUT_DIR}\n\n请在该目录下按厂家建子文件夹。")
            return

        suppliers = sorted([
            d for d in os.listdir(INPUT_DIR)
            if os.path.isdir(os.path.join(INPUT_DIR, d))
        ])

        # 清空旧勾选框
        for w in self.check_frame.winfo_children():
            w.destroy()
        self._check_vars.clear()

        if not suppliers:
            ttk.Label(self.check_frame, text="(暂无厂家子文件夹)", foreground="#888").pack(anchor="w")
            self._suppliers = []
            return

        # 垂直排列 Checkbutton
        for i, s in enumerate(suppliers):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.check_frame, text=s, variable=var)
            cb.grid(row=i, column=0, sticky="w", padx=4, pady=2)
            self._check_vars[s] = var

        self._suppliers = suppliers
        self._log(f"📋 发现 {len(suppliers)} 个厂家: {', '.join(suppliers)}")

    def _select_all(self):
        for v in self._check_vars.values():
            v.set(True)

    def _clear_all(self):
        for v in self._check_vars.values():
            v.set(False)

    def _get_selected(self):
        return [s for s, v in self._check_vars.items() if v.get()]

    # ---------------- 按钮回调 ----------------
    def _on_start(self):
        if self._running:
            return

        selected = self._get_selected()
        if not selected:
            messagebox.showwarning("提示", "请至少选一个厂家")
            return

        # 验证码：先查缓存（一天只输一次），缓存命中直接跳过
        if not self.auth_ok:
            from auth import _is_cached_today
            if _is_cached_today():
                self.auth_ok = True
                self.auth_label.configure(text="🔓 已验证（今日缓存）", foreground="green")
            elif not self._do_auth():
                return
            else:
                self.auth_label.configure(text="🔓 已验证", foreground="green")

        self._running = True
        self.start_btn.configure(state="disabled", text="⏳ 处理中...")

        # 子线程跑处理
        t = threading.Thread(target=self._run_processing, args=(selected,), daemon=True)
        t.start()

    def _do_auth(self):
        """弹窗让用户输验证码。返回 True 表示通过。"""
        from auth import verify_with_grace, _save_cache, gen_code, date

        dlg = tk.Toplevel(self.root)
        dlg.title("授权验证")
        dlg.geometry("380x200")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        tk.Label(dlg, text=f"今日日期: {date.today().strftime('%Y-%m-%d')}", font=("", 11)).pack(pady=(16, 4))
        tk.Label(dlg, text="请输入今日验证码：").pack(pady=4)

        entry = tk.Entry(dlg, font=("", 14), width=12, justify="center")
        entry.pack(pady=8)
        entry.focus_set()

        result = {"ok": False}

        def on_ok(_evt=None):
            code = entry.get().strip()
            if verify_with_grace(code, grace_days=1):
                _save_cache()  # 缓存今日验证结果，一天只输一次
                result["ok"] = True
                dlg.destroy()
            else:
                tk.Label(dlg, text="验证码错误，请重试", foreground="red").pack(pady=2)
                entry.delete(0, "end")
                entry.focus_set()

        def on_cancel():
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=8)
        ttk.Button(btn_frame, text="确定", width=10, command=on_ok).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="取消", width=10, command=on_cancel).pack(side="left", padx=8)

        entry.bind("<Return>", on_ok)
        dlg.protocol("WM_DELETE_WINDOW", on_cancel)
        dlg.wait_window()

        return result["ok"]

    def _open_input_dir(self):
        try:
            from config import INPUT_DIR
            self._open_folder(INPUT_DIR)
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def _open_output_dir(self):
        try:
            from config import OUTPUT_DIR
            self._open_folder(OUTPUT_DIR)
        except Exception as e:
            messagebox.showerror("错误", str(e))

    @staticmethod
    def _open_folder(path):
        if not os.path.exists(path):
            os.makedirs(path, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    # ---------------- 处理线程 ----------------
    def _run_processing(self, selected_suppliers):
        """子线程里跑处理流程。"""
        try:
            # 重定向 stdout/stderr 到 GUI 日志队列
            old_stdout, old_stderr = sys.stdout, sys.stderr
            sys.stdout = _GuiStdout(self._log_q)
            sys.stderr = _GuiStderr(self._log_q)

            try:
                from excel_processor import process_all, OUTPUT_DIR
                from yc_processor import process_yuanchuang_all

                self._log(f"\n{'═' * 50}")
                self._log(f"🚀 开始处理，选中厂家: {', '.join(selected_suppliers)}")
                self._log(f"{'═' * 50}")

                file_results = process_all(suppliers=selected_suppliers)
                yc_results = process_yuanchuang_all(suppliers=selected_suppliers)

                all_results = (file_results or []) + (yc_results or [])

                if not all_results:
                    self._log("\n⚠ 没有可处理的文件")
                else:
                    self._print_summary(all_results, OUTPUT_DIR)

                self._log(f"\n✅ 本轮处理完成！可继续选厂家跑下一轮。")

            finally:
                sys.stdout = old_stdout
                sys.stderr = old_stderr

        except Exception as e:
            self._log(f"\n❌ 处理出错: {e}")
            import traceback
            for line in traceback.format_exc().splitlines():
                self._log(line)

        finally:
            # 回到主线程收尾
            self.root.after(0, self._finish_processing)

    def _finish_processing(self):
        self._running = False
        self.start_btn.configure(state="normal", text="▶ 开始处理")

    def _print_summary(self, results, output_dir):
        """GUI 版汇总日志。"""
        self._log("")
        self._log("=" * 60)
        self._log("📊 处理汇总")
        self._log("=" * 60)

        supplier_totals = {}
        for r in results:
            supplier = r.get("supplier", "未知")
            supplier_totals.setdefault(supplier, {"files": 0, "rows": 0, "imgs": 0})
            supplier_totals[supplier]["files"] += 1
            supplier_totals[supplier]["rows"] += r.get("rows", 0)
            supplier_totals[supplier]["imgs"] += r.get("image_count", 0)

        for supplier, t in supplier_totals.items():
            self._log(f"  [{supplier}] {t['files']}文件 / {t['rows']}行 / {t['imgs']}图片")

        total = sum(v["files"] for v in supplier_totals.values())
        self._log(f"  ── 合计 {total} 个文件 ──")
        self._log(f"📁 输出目录: {output_dir}")
        self._log("=" * 60)

    # ---------------- 日志 ----------------
    def _drain_log_queue(self):
        """主线程定时从 Queue 取日志，刷到 Text 控件。"""
        try:
            while True:
                line = self._log_q.get_nowait()
                self._append_log(line)
        except queue.Empty:
            pass
        self.root.after(80, self._drain_log_queue)

    def _append_log(self, line):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log(self, msg):
        """子线程安全地发日志（进 Queue）。"""
        self._log_q.put(msg)

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")


# =============================================================
# 入口
# =============================================================
def launch_gui():
    root = tk.Tk()
    # Windows 下用系统默认主题，外观和原生一致
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    App(root)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
