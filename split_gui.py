# -*- coding: utf-8 -*-
"""
供应商 + 门店 配置管理 + 一键拆分  图形界面
==========================================
- 供应商列表：新增 / 编辑 / 删除（开票状态、折扣、收款信息、不开票收款信息）
- 门店列表：新增 / 编辑 / 删除（归属三档：单独拆分 / 总部其他 / 分公司；口径：开票 / 不开票）
- 底部一键「执行拆分」，日志实时输出
- 配置统一保存到 supplier_config.json（与脚本/exe 同目录）
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


class StoreDialog(tk.Toplevel):
    """新增/编辑门店对话框：编号 + 归属（三档）+ 口径（开票/不开票）"""

    def __init__(self, parent, title, code="", data=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.grab_set()

        data = data or {}
        pad = {"padx": 8, "pady": 4}

        row = 0
        ttk.Label(self, text="门店编号").grid(row=row, column=0, sticky="e", **pad)
        self.var_code = tk.StringVar(value=code)
        self.ent_code = ttk.Entry(self, textvariable=self.var_code, width=30)
        self.ent_code.grid(row=row, column=1, sticky="w", **pad)
        if code:  # 编辑已有门店时编号锁定
            self.ent_code.config(state="disabled")
        row += 1

        ttk.Label(self, text="门店归属").grid(row=row, column=0, sticky="e", **pad)
        self.var_group = tk.StringVar(
            value=data.get("group", supplier_config.STORE_GROUP_HQ))
        self.cmb_group = ttk.Combobox(self, textvariable=self.var_group,
                                      values=supplier_config.STORE_GROUPS,
                                      state="readonly", width=27)
        self.cmb_group.grid(row=row, column=1, sticky="w", **pad)
        self.cmb_group.bind("<<ComboboxSelected>>", self._on_group)
        row += 1

        ttk.Label(self, text="金额口径").grid(row=row, column=0, sticky="e", **pad)
        self.var_invoice = tk.StringVar(
            value=data.get("invoice", supplier_config.INVOICE_ON))
        self.cmb_invoice = ttk.Combobox(self, textvariable=self.var_invoice,
                                        values=supplier_config.INVOICE_MODES,
                                        state="readonly", width=27)
        self.cmb_invoice.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        self.lbl_hint = ttk.Label(self, text="", foreground="#888888")
        self.lbl_hint.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        bf = ttk.Frame(self)
        bf.grid(row=row, column=0, columnspan=2, pady=10)
        ttk.Button(bf, text="保存", command=self._ok).pack(side="left", padx=6)
        ttk.Button(bf, text="取消", command=self.destroy).pack(side="left", padx=6)

        self._on_group()

    def _on_group(self, *_):
        g = self.var_group.get()
        if g == supplier_config.STORE_GROUP_SPLIT:
            self.cmb_invoice.config(state="readonly")
            self.lbl_hint.config(text="单独出表：开票=单店口径，不开票=45/48口径")
        else:
            self.cmb_invoice.config(state="disabled")
            self.lbl_hint.config(text="并入汇总表，金额口径固定为不开票")

    def _ok(self):
        code = self.var_code.get().strip()
        if not code:
            messagebox.showwarning("提示", "门店编号不能为空", parent=self)
            return
        self.result = {
            "code": code,
            "group": self.var_group.get(),
            "invoice": self.var_invoice.get(),
        }
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("供应商对账单拆分工具")
        self.geometry("900x660")
        self.minsize(760, 560)

        self.logq = queue.Queue()
        self._build_ui()
        self._refresh_suppliers()
        self._refresh_stores()
        self.after(100, self._poll_log)

    # ---------- UI ----------
    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="执行拆分", command=self._run_split).pack(side="right", padx=4)
        ttk.Button(top, text="打开输出目录", command=self._open_output).pack(side="right", padx=4)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        # ---- 供应商页 ----
        sup = ttk.Frame(nb, padding=6)
        nb.add(sup, text="供应商")
        sup_top = ttk.Frame(sup)
        sup_top.pack(fill="x")
        ttk.Button(sup_top, text="新增供应商", command=self._add_supplier).pack(side="left", padx=4)
        ttk.Button(sup_top, text="编辑", command=self._edit_supplier).pack(side="left", padx=4)
        ttk.Button(sup_top, text="删除", command=self._delete_supplier).pack(side="left", padx=4)
        ttk.Button(sup_top, text="刷新", command=self._refresh_suppliers).pack(side="left", padx=4)

        cols = ("name", "status", "discount", "bank")
        self.tree = ttk.Treeview(sup, columns=cols, show="headings")
        self.tree.heading("name", text="供应商")
        self.tree.heading("status", text="开票状态")
        self.tree.heading("discount", text="折扣")
        self.tree.heading("bank", text="收款信息（开票）")
        self.tree.column("name", width=140, anchor="w")
        self.tree.column("status", width=100, anchor="center")
        self.tree.column("discount", width=70, anchor="center")
        self.tree.column("bank", width=520, anchor="w")
        vsb = ttk.Scrollbar(sup, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True, pady=(4, 0))
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self._edit_supplier())

        # ---- 门店页 ----
        sto = ttk.Frame(nb, padding=6)
        nb.add(sto, text="门店")
        sto_top = ttk.Frame(sto)
        sto_top.pack(fill="x")
        ttk.Button(sto_top, text="新增门店", command=self._add_store).pack(side="left", padx=4)
        ttk.Button(sto_top, text="编辑", command=self._edit_store).pack(side="left", padx=4)
        ttk.Button(sto_top, text="删除", command=self._delete_store).pack(side="left", padx=4)
        ttk.Button(sto_top, text="刷新", command=self._refresh_stores).pack(side="left", padx=4)

        scols = ("code", "group", "invoice")
        self.stree = ttk.Treeview(sto, columns=scols, show="headings")
        self.stree.heading("code", text="门店编号")
        self.stree.heading("group", text="门店归属")
        self.stree.heading("invoice", text="金额口径")
        self.stree.column("code", width=160, anchor="w")
        self.stree.column("group", width=200, anchor="center")
        self.stree.column("invoice", width=160, anchor="center")
        svsb = ttk.Scrollbar(sto, orient="vertical", command=self.stree.yview)
        self.stree.configure(yscrollcommand=svsb.set)
        self.stree.pack(fill="both", expand=True, pady=(4, 0))
        svsb.pack(side="right", fill="y")
        self.stree.bind("<Double-1>", lambda e: self._edit_store())

        # 日志区
        logf = ttk.LabelFrame(self, text="运行日志", padding=4)
        logf.pack(fill="x", padx=8, pady=8)
        self.log = tk.Text(logf, height=10, state="disabled", wrap="word")
        self.log.pack(fill="x")
        self.run_btn = None

    # ---------- 供应商 ----------
    def _refresh_suppliers(self):
        self.tree.delete(*self.tree.get_children())
        cfg = supplier_config.load_config(force_reload=True)
        for name, d in sorted(cfg.get("suppliers", {}).items()):
            bank = (d.get("bank") or "").replace("\n", " ")
            self.tree.insert("", "end", iid=name, values=(
                name, d.get("invoice_status", ""),
                discount_to_label(d.get("discount", 1.00)), bank))

    def _add_supplier(self):
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
            self._refresh_suppliers()

    def _edit_supplier(self):
        name = self._sel("tree")
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
            self._refresh_suppliers()

    def _delete_supplier(self):
        name = self._sel("tree")
        if not name:
            messagebox.showinfo("提示", "请先选择一个供应商")
            return
        if not messagebox.askyesno("确认", f"确定删除供应商「{name}」？"):
            return
        cfg = supplier_config.load_config()
        cfg.get("suppliers", {}).pop(name, None)
        supplier_config.save_config(cfg)
        self._refresh_suppliers()

    # ---------- 门店 ----------
    def _refresh_stores(self):
        self.stree.delete(*self.stree.get_children())
        cfg = supplier_config.load_config(force_reload=True)
        for code, d in sorted(cfg.get("stores", {}).items()):
            self.stree.insert("", "end", iid=code, values=(
                code, d.get("group", ""), d.get("invoice", "")))

    def _add_store(self):
        dlg = StoreDialog(self, "新增门店")
        self.wait_window(dlg)
        if dlg.result:
            supplier_config.upsert_store(
                dlg.result["code"], dlg.result["group"], dlg.result["invoice"])
            self._refresh_stores()

    def _edit_store(self):
        code = self._sel("stree")
        if not code:
            messagebox.showinfo("提示", "请先选择一个门店")
            return
        data = supplier_config.get_store(code) or {}
        dlg = StoreDialog(self, "编辑门店", code=code, data=data)
        self.wait_window(dlg)
        if dlg.result:
            supplier_config.upsert_store(
                dlg.result["code"], dlg.result["group"], dlg.result["invoice"])
            self._refresh_stores()

    def _delete_store(self):
        code = self._sel("stree")
        if not code:
            messagebox.showinfo("提示", "请先选择一个门店")
            return
        if not messagebox.askyesno("确认", f"确定删除门店「{code}」？"):
            return
        supplier_config.delete_store(code)
        self._refresh_stores()

    def _sel(self, tree_key):
        tree = self.tree if tree_key == "tree" else self.stree
        sel = tree.selection()
        return sel[0] if sel else None

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