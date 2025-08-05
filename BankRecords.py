#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
交易流水分析 GUI – 纯代码像素背景 (no-alpha colors)
Author: 温岭纪委六室 单柳昊
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, builtins, warnings, datetime, re
from pathlib import Path
from typing import Optional
import pandas as pd, numpy as np
from chinese_calendar import is_holiday, is_workday

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
OUT_DIR: Optional[Path] = None       # 运行期赋值

# ========= 业务函数（与之前版本相同，略） =========
# ……（为了篇幅，此处省略业务函数，可以直接粘贴你原先的 merge_txn_by_name、
#     analysis_txn、make_partner_summary、save_df_auto_width 等函数）……
# ==============================================

# ------------------------------------------------
# 生成浅蓝→白渐变背景
def generate_gradient_photo(w: int, h: int) -> tk.PhotoImage:
    img = tk.PhotoImage(width=w, height=h)
    for y in range(h):
        ratio = y / h
        r = int(0xde + (0xff - 0xde) * ratio)  # 0xDE → 0xFF
        g = int(0xf7 + (0xff - 0xf7) * ratio)  # 0xF7 → 0xFF
        b = 0xff
        img.put(f"#{r:02x}{g:02x}{b:02x}", to=(0, y, w, y + 1))
    return img

# ------------------------------------------------
def create_gui():
    root = tk.Tk()
    root.title("温岭纪委交易流水分析工具")
    root.minsize(780, 560)

    # 背景
    bg_photo = generate_gradient_photo(root.winfo_screenwidth(), root.winfo_screenheight())
    canvas = tk.Canvas(root, highlightthickness=0)
    canvas.pack(fill="both", expand=True)
    canvas.create_image(0, 0, image=bg_photo, anchor="nw")

    # 主容器（半透明改为纯白）
    main = tk.Frame(canvas, bg="#ffffff")
    canvas.create_window(0, 0, window=main, anchor="nw")

    # ttk 主题
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TLabel", font=("微软雅黑", 11), background="#ffffff")
    style.configure("TButton", font=("微软雅黑", 10, "bold"), padding=5)
    style.configure("TLabelframe", background="#ffffff")
    style.configure("TLabelframe.Label", font=("微软雅黑", 11, "bold"))

    # ===== 标题 =====
    tk.Label(main, text="温岭纪委交易流水分析工具",
             font=("微软雅黑", 20, "bold"), bg="#ffffff").grid(row=0, column=0, columnspan=3,
                                                               pady=(18, 4), padx=20, sticky="w")
    tk.Label(main, text="© 温岭纪委六室 单柳昊",
             font=("微软雅黑", 9), fg="#555", bg="#ffffff").grid(row=1, column=0, columnspan=3,
                                                                sticky="w", padx=22)

    # ===== 输入区 =====
    frm = ttk.Labelframe(main, text="输入参数")
    frm.grid(row=2, column=0, columnspan=3, padx=20, pady=10, sticky="we")
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text="工作目录:").grid(row=0, column=0, sticky="e", pady=6, padx=4)
    path_var = tk.StringVar(value=str(Path.home()))
    ttk.Entry(frm, textvariable=path_var).grid(row=0, column=1, sticky="we", pady=6)
    ttk.Button(frm, text="浏览", width=8,
               command=lambda: path_var.set(filedialog.askdirectory() or path_var.get())
               ).grid(row=0, column=2, padx=6, pady=6)

    ttk.Label(frm, text="账户姓名:").grid(row=1, column=0, sticky="e", pady=6, padx=4)
    name_var = tk.StringVar()
    ttk.Entry(frm, textvariable=name_var, width=25).grid(row=1, column=1, sticky="w", pady=6)

    # ===== 日志区 =====
    log_f = ttk.Labelframe(main, text="运行日志")
    log_f.grid(row=4, column=0, columnspan=3, padx=20, pady=(0, 20), sticky="nsew")
    log_f.rowconfigure(0, weight=1)
    log_f.columnconfigure(0, weight=1)
    log_box = tk.Text(log_f, height=16, bg="#f8f8f8", state="disabled")
    log_box.grid(row=0, column=0, sticky="nsew")

    # ===== 开始按钮 =====
    start_btn = ttk.Button(main, text="开始分析", width=18)
    start_btn.grid(row=3, column=0, columnspan=3, pady=8)

    # main 自适应
    root.update_idletasks()
    root.bind("<Configure>", lambda e: main.config(width=e.width, height=e.height))

    # -------- 日志函数 --------
    def log(msg: str):
        log_box.config(state="normal")
        log_box.insert("end", f"{datetime.datetime.now():%H:%M:%S}  {msg}\n")
        log_box.config(state="disabled")
        log_box.see("end")

    # -------- 核心流程 --------
    def run(path: str, holder: str):
        global OUT_DIR
        OUT_DIR = Path(path).expanduser().resolve() / f"{holder}-分析"
        OUT_DIR.mkdir(parents=True, exist_ok=True)

        _orig_print = builtins.print
        builtins.print = lambda *a, **kw: log(" ".join(map(str, a)))
        start_btn.config(state="disabled")
        try:
            all_txn = merge_txn_by_name(path, holder)
            analysis_txn(all_txn)
            make_partner_summary(all_txn)
            messagebox.showinfo("完成", f"分析已完成！\n结果保存在:\n{OUT_DIR}")
        except Exception as ex:
            messagebox.showerror("错误", str(ex))
        finally:
            builtins.print = _orig_print
            start_btn.config(state="normal")

    start_btn.configure(command=lambda: threading.Thread(
        target=run, args=(path_var.get(), name_var.get()), daemon=True).start())

    root.mainloop()


if __name__ == "__main__":
    create_gui()
