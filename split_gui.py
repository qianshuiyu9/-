# -*- coding: utf-8 -*-
"""
供应商配置管理 + 一键拆分  图形界面
==================================
- 供应商列表：新增 / 编辑 / 删除
- 每个供应商维护：开票状态（三档）、折扣、收款信息、不开票收款信息
- 底部一键「执行拆分」，日志实时输出
- 配置保存到 supplier_config.json（与脚本/exe 同目录）
"""
import queue
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import supplier_config


DISCOUNT_PRESETS = [
    ("不打折", 1.00),
    ("99折", 0.99),
    ("98折", 0.98),
    ("97折", 0.97),
    ("96折", 0.96),
    ("95折", 0.95),
]


def discount_to_label(d):
    for label, val in DISCOUNT_PRESETS:
        if abs(float(d) - val) < 1e-9:
            return label
    return f"{float(d):g}"


class SupplierDialog(tk.Toplevel):
    """新增/编辑供应商对话框"""

    def __init__(self, parent, title, name="", data=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.grab_set()

        data = data or {}
        pad = {"padx": 8, "pady": 4}

        row = 0
        ttk.Label(self, text="供应商名称").grid(row=row, column=0, sticky="e", **pad)
        self.var_name = tk.StringVar(value=name)
        self.ent_name = ttk.Entry(self, textvariable=self.var_name, width=32)
        self.ent_name.grid(row=row, column=1, sticky="w", **pad)
        if name:  # 编辑已有供应商时名称锁定
            self.ent_name.config(state="disabled")
        row += 1

        ttk.Label(self, text="开票状态").grid(row=row, column=0, sticky="e", **pad)
        self.var_status = tk.StringVar(
            value=data.get("invoice_status", supplier_config.STATUS_NO_PLUS))
        self.cmb_status = ttk.Combobox(self, textvariable=self.var_status,
                                       values=supplier_config.STATUSES,
                                       state="readonly", width=29)
        self.cmb_status.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        ttk.Label(self, text="折扣").grid(row=row, column=0, sticky="e", **pad)
        label = discount_to_label(data.get("discount", 1.00))
        self.var_discount = tk.StringVar(value=label)
        self.cmb_discount = ttk.Combobox(self, textvariable=self.var_discount,
                                         values=[l for l, _ in DISCOUNT_PRESETS],
                                         state="readonly", width=29)
        self.cmb_discount.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        ttk.Label(self, text="收款信息（开票）").grid(row=row, column=0, sticky="ne", **pad)
        self.txt_bank = tk.Text(self, width=36, height=5)
        self.txt_bank.grid(row=row, column=1, sticky="w", **pad)
        self.txt_bank.insert("1.0", data.get("bank", ""))
        row += 1

        ttk.Label(self, text="收款信息（不开票）").grid(row=row, column=0, sticky="ne", **pad)
        self.txt_nobank = tk.Text(self, width=36, height=5)
        self.txt_nobank.grid(row=row, column=1, sticky="w", **pad)
        self.txt_nobank.insert("1.0", data.get("no_bank", ""))
        row += 1

        bf = ttk.Frame(self)
        bf.grid(row=row, column=0, columnspan=2, pady=10)
        ttk.Button(bf, text="保存", command=self._ok).pack(side="left", padx=6)
        ttk.Button(bf, text="取消", command=self.destroy).pack(side="left", padx=6)

    def _ok(self):
        name = self.var_name.get().strip()
        if not name:
            messagebox.showwarning("提示", "供应商名称不能为空", parent=self)
            return
        label = self.var_discount.get()
        discount = 1.00
        for l, v in DISCOUNT_PRESETS:
            if l == label:
                discount = v
                break
        self.result = {
            "name": name,
            "invoice_status": self.var_status.get(),
            "discount": discount,
            "bank": self.txt_bank.get("1.0", "end").strip(),
            "no_bank": self.txt_nobank.get("1.0", "end").strip(),
        }
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("供应商对账单拆分工具")
        self.geometry("860x620")
        self.minsize(720, 520)

        self.logq = queue.Queue()
        self._build_ui()
        self._refresh_list()
        self.after(100, self._poll_log)

    # ---------- UI ----------
    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Button(top, text="新增供应商", command=self._add).pack(side="left", padx=4)
        ttk.Button(top, text="编辑", command=self._edit).pack(side="left", padx=4)
        ttk.Button(top, text="删除", command=self._delete).pack(side="left", padx=4)
        ttk.Button(top, text="刷新", command=self._refresh_list).pack(side="left", padx=4)

        ttk.Button(top, text="执行拆分", command=self._run_split).pack(side="right", padx=4)
        ttk.Button(top, text="打开输出目录", command=self._open_output).pack(side="right", padx=4)

        # 供应商列表
        cols = ("name", "status", "discount", "bank")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        self.tree.heading("name", text="供应商")
        self.tree.heading("status", text="开票状态")
        self.tree.heading("discount", text="折扣")
        self.tree.heading("bank", text="收款信息（开票）")
        self.tree.column("name", width=140, anchor="w")
        self.tree.column("status", width=100, anchor="center")
        self.tree.column("discount", width=70, anchor="center")
        self.tree.column("bank", width=520, anchor="w")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True, padx=8, pady=(4, 0))
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self._edit())

        # 日志区
        logf = ttk.LabelFrame(self, text="运行日志", padding=4)
        logf.pack(fill="x", padx=8, pady=8)
        self.log = tk.Text(logf, height=10, state="disabled", wrap="word")
        self.log.pack(fill="x")
        self.run_btn = None

    # ---------- 供应商列表 ----------
    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        cfg = supplier_config.load_config(force_reload=True)
        for name, d in sorted(cfg.get("suppliers", {}).items()):
            bank = (d.get("bank") or "").replace("\n", " ")
            self.tree.insert("", "end", iid=name, values=(
                name, d.get("invoice_status", ""),
                discount_to_label(d.get("discount", 1.00)), bank))

    def _selected(self):
        sel = self.tree.selection()
        return sel[0] if sel else None

    def _add(self):
        dlg = SupplierDialog(self, "新增供应商")
        self.wait_window(dlg)
        if dlg.result:
            cfg = supplier_config.load_config()
            cfg.setdefault("suppliers", {})[dlg.result["name"]] = {
                "invoice_status": dlg.result["invoice_status"],
                "discount": dlg.result["discount"],
                "bank": dlg.result["bank"],
                "no_bank": dlg.result["no_bank"],
            }
            supplier_config.save_config(cfg)
            self._refresh_list()

    def _edit(self):
        name = self._selected()
        if not name:
            messagebox.showinfo("提示", "请先选择一个供应商")
            return
        data = supplier_config.get_supplier(name) or {}
        dlg = SupplierDialog(self, "编辑供应商", name=name, data=data)
        self.wait_window(dlg)
        if dlg.result:
            cfg = supplier_config.load_config()
            cfg.setdefault("suppliers", {})[name] = {
                "invoice_status": dlg.result["invoice_status"],
                "discount": dlg.result["discount"],
                "bank": dlg.result["bank"],
                "no_bank": dlg.result["no_bank"],
            }
            supplier_config.save_config(cfg)
            self._refresh_list()

    def _delete(self):
        name = self._selected()
        if not name:
            messagebox.showinfo("提示", "请先选择一个供应商")
            return
        if not messagebox.askyesno("确认", f"确定删除供应商「{name}」？"):
            return
        cfg = supplier_config.load_config()
        cfg.get("suppliers", {}).pop(name, None)
        supplier_config.save_config(cfg)
        self._refresh_list()

    # ---------- 拆分 ----------
    def _open_output(self):
        import split_reconcile_by_store as engine
        import os
        os.startfile(engine.OUTPUT_DIR)

    def _run_split(self):
        for child in self.winfo_children():
            if isinstance(child, ttk.Button) and child.cget("text") == "执行拆分":
                self.run_btn = child
                break
        if self.run_btn:
            self.run_btn.config(state="disabled")
        self._log("开始拆分...\n")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        import io
        import contextlib
        import split_reconcile_by_store as engine

        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf):
                engine.run_split(engine.INPUT_DIR, engine.OUTPUT_DIR)
        except Exception:
            import traceback
            traceback.print_exc(file=buf)
        finally:
            self.logq.put(("done", buf.getvalue()))

    # ---------- 日志 ----------
    def _log(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.config(state="disabled")

    def _poll_log(self):
        try:
            while True:
                kind, text = self.logq.get_nowait()
                self._log(text)
                if kind == "done":
                    if self.run_btn:
                        self.run_btn.config(state="normal")
                        self.run_btn = None
                    self._log("\n—— 拆分完成 ——\n")
        except queue.Empty:
            pass
        self.after(100, self._poll_log)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()