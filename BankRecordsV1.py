#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
交易流水分析 GUI
Author: 温岭纪委六室单柳昊
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import warnings
import builtins
import sys
from pathlib import Path
import re
import pandas as pd
import numpy as np
import datetime
from typing import Optional
from chinese_calendar import is_holiday, is_workday

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUT_DIR: Optional[Path] = None

full_ts_pat = re.compile(r"\d{4}-\d{2}-\d{2}-\d{2}\.\d{2}\.\d{2}\.\d+")


def _normalize_time(t: str, is_old: bool) -> str:
    if not t:
        return ""
    if "." in t:
        t = t.replace(".", ":")
        t = re.sub(r":\d{1,6}$", "", t)
    if re.fullmatch(r"\d{6}", t):
        t = f"{t[:2]}:{t[2:4]}:{t[4:]}"
    elif re.fullmatch(r"\d{4}", t):
        t = f"{t[:2]}:{t[2:]}:00"
    if is_old and re.fullmatch(r"\d{6}", t.replace(":", "")):
        t_num = t.replace(":", "")
        t = f"{t_num[:2]}:{t_num[2:4]}:{t_num[4:]}"
    return t


def save_df_auto_width(
    df: pd.DataFrame,
    filename: Path | str,
    sheet_name: str = "Sheet1",
    index: bool = False,
    engine: str = "xlsxwriter",
    min_width: int = 6,
    max_width: int = 50,
):
    """自动列宽保存 Excel；filename 可以是相对名，自动写入 OUT_DIR"""
    if OUT_DIR is not None:
        filename = OUT_DIR / filename
    filename = Path(filename).with_suffix(".xlsx")

    df = df.replace(np.nan, "")
    if engine not in {"xlsxwriter", "openpyxl"}:
        raise ValueError("engine 只能是 'xlsxwriter' 或 'openpyxl'")
    if engine == "xlsxwriter":
        with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=index)
            ws = writer.sheets[sheet_name]
            for idx, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len, na_action="ignore").max(),
                    len(str(col)),
                )
                width = max(min_width, min(max_len + 2, max_width))
                ws.set_column(idx, idx, width)
    else:  # openpyxl
        df.to_excel(filename, sheet_name=sheet_name, index=index, engine="openpyxl")
        from openpyxl import load_workbook

        wb = load_workbook(filename)
        ws = wb[sheet_name]
        for col_cells in ws.columns:
            max_len = 0
            for cell in col_cells:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            letter = col_cells[0].column_letter
            width = max(min_width, min(max_len + 2, max_width)) + 5
            ws.column_dimensions[letter].width = width
        wb.save(filename)


def str_to_weekday(date_val) -> str:
    """
    将任意类型的日期时间值映射为中文星期。
    无法解析时返回空字符串。
    """
    dt = pd.to_datetime(date_val, errors="coerce")
    if pd.isna(dt):
        return ""
    cn_weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    return cn_weekdays[dt.weekday()]


def holiday_status(date_val) -> str:
    """
    判断给定日期是节假日 / 周末 / 工作日。
    无法解析时返回空字符串。
    """
    dt = pd.to_datetime(date_val, errors="coerce")
    if pd.isna(dt):
        return ""
    date_obj = dt.date()
    if is_holiday(date_obj):
        return "节假日"
    elif is_workday(date_obj):
        return "工作日"
    else:
        return "周末"


def merge_txn_by_name(root_dir: str, name: str) -> pd.DataFrame:
    root = Path(root_dir).expanduser().resolve()
    patterns = [
        f"{name}-*-交易流水.xlsx",
        f"{name}-*-交易流水.xls",
        "*老农商银行*.xlsx",
        "*老农商银行*.xls",
        "*新农商银行*.xlsx",
        "*新农商银行*.xls",
    ]
    candidates = []
    for pat in patterns:
        candidates.extend(root.rglob(pat))

    china_regex = re.compile(rf"^{re.escape(name)}-[^-]+-交易流水\.xls[x]?$", re.I)
    rc_regex = re.compile(r".*(新农商银行|老农商银行).*\.xls[x]?$", re.I)

    china_files = [p for p in candidates if china_regex.match(p.name)]
    rc_files = [p for p in candidates if rc_regex.match(p.name)]

    print(f"✅ 找到 {len(china_files)} 份 {name} 的线上银行流水")
    print(f"✅ 找到 {len(rc_files)} 份 新/老农商银行流水")

    dfs = []
    for p in china_files:
        try:
            df = pd.read_excel(
                p,
                dtype={
                    "查询卡号": str,
                    "查询账户": str,
                    "交易对方证件号码": str,
                    "本方账号": str,
                    "本方卡号": str,
                },
            )
            df["来源文件"] = p.name
            dfs.append(df)
        except Exception as e:
            print(f"❌ 读取 {p.name} 失败: {e}")

    TEMPLATE_COLS = [
        "序号",
        "查询对象",
        "反馈单位",
        "查询项",
        "查询账户",
        "查询卡号",
        "交易类型",
        "借贷标志",
        "币种",
        "交易金额",
        "账户余额",
        "交易时间",
        "交易流水号",
        "本方账号",
        "本方卡号",
        "交易对方姓名",
        "交易对方账户",
        "交易对方卡号",
        "交易对方证件号码",
        "交易对手余额",
        "交易对方账号开户行",
        "交易摘要",
        "交易网点名称",
        "交易网点代码",
        "日志号",
        "传票号",
        "凭证种类",
        "凭证号",
        "现金标志",
        "终端号",
        "交易是否成功",
        "交易发生地",
        "商户名称",
        "商户号",
        "IP地址",
        "MAC",
        "交易柜员号",
        "备注",
    ]

    def _header_row(path: Path) -> int:
        raw = pd.read_excel(path, header=None, nrows=15)
        for i, r in raw.iterrows():
            if "交易日期" in r.values:
                return i
        return 0

    def _read_rc_raw(path: Path) -> pd.DataFrame:
        return pd.read_excel(path, header=_header_row(path))

    def rc_to_template(df_raw: pd.DataFrame, holder: str, is_old: bool) -> pd.DataFrame:
        out = pd.DataFrame(columns=TEMPLATE_COLS)
        out["本方账号"] = out["查询账户"] = df_raw.get("账号")
        out["交易金额"] = df_raw.get("发生额") if is_old else df_raw.get("交易金额")
        out["账户余额"] = df_raw.get("余额") if is_old else df_raw.get("交易余额")
        out["反馈单位"] = "老农商银行" if is_old else "新农商银行"
        date_series = df_raw.get("交易日期").astype(str).fillna("")
        time_series = df_raw.get("交易时间").astype(str).fillna("")

        def _parse_datetime(d, t):
            t = t.strip()
            if full_ts_pat.fullmatch(t):
                dt = pd.to_datetime(t, format="%Y-%m-%d-%H.%M.%S.%f", errors="coerce")
            else:
                t_norm = _normalize_time(t, is_old)
                dt = pd.to_datetime(f"{d} {t_norm}".strip(), errors="coerce")
            return dt.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(dt) else ""

        out["交易时间"] = [_parse_datetime(d, t) for d, t in zip(date_series, time_series)]
        flag = df_raw.get("借贷标志")
        out["借贷标志"] = flag.map({"1": "进", "2": "出"}) if is_old else flag
        out["币种"] = "CNY"
        out["查询对象"] = holder
        out["交易对方姓名"] = " " if is_old else df_raw.get("对方姓名")
        out["交易对方账户"] = " " if is_old else df_raw.get("对方账号")
        out["交易摘要"] = df_raw.get("备注") if is_old else df_raw.get("交易码转义")
        return out

    for p in rc_files:
        try:
            holder = p.stem.split("-")[0]
            raw = _read_rc_raw(p)
            df = rc_to_template(raw, holder, is_old=("老农商银行" in p.name))
            df["来源文件"] = p.name
            dfs.append(df)
        except Exception as e:
            print(f"❌ 读取/转换 {p.name} 失败: {e}")

    if not dfs:
        raise RuntimeError("⚠️ 未找到可合并的文件")

    all_txn = pd.concat(dfs, ignore_index=True)
    ts = pd.to_datetime(all_txn["交易时间"], errors="coerce")
    all_txn.insert(0, "__ts__", ts)
    all_txn.sort_values("__ts__", inplace=True)
    all_txn["序号"] = range(1, len(all_txn) + 1)
    all_txn.drop(columns="__ts__", inplace=True)
    all_txn["星期"] = all_txn["交易时间"].apply(str_to_weekday)
    all_txn["节假日"] = all_txn["交易时间"].apply(holiday_status)
    bins = [-np.inf, 2000, 5000, 20000, 50000, np.inf]
    labels = ["2000以下", "2000-5000", "5000-20000", "20000-50000", "50000以上"]
    all_txn["金额区间"] = pd.cut(
        all_txn["交易金额"], bins=bins, labels=labels, right=False, include_lowest=True
    )
    out_path = f"{name}-合并交易流水.xlsx"
    save_df_auto_width(all_txn, filename=out_path, index=False, engine="openpyxl")
    print(f"✅ 合并完成，已导出 {out_path}")
    return all_txn


def make_partner_summary(df: pd.DataFrame) -> None:
    df = df.copy()
    df["交易金额"] = pd.to_numeric(df["交易金额"], errors="coerce").fillna(0)
    df["is_in"] = df["借贷标志"] == "进"
    df["is_out"] = df["借贷标志"] == "出"
    df["abs_amt"] = df["交易金额"].abs()
    df["signed_amt"] = df["交易金额"] * df["is_in"].map({True: 1, False: -1})
    df["in_amt"] = df["交易金额"].where(df["is_in"], 0)
    df["out_amt"] = df["交易金额"].where(df["is_out"], 0)
    df["gt10k"] = (df["abs_amt"] >= 10_000).astype(int)
    grp_cols = ["查询对象", "交易对方姓名"]
    summary = (
        df.groupby(grp_cols, dropna=False)
        .agg(
            交易次数=("交易金额", "size"),
            交易金额=("abs_amt", "sum"),
            万元以上交易次数=("gt10k", "sum"),
            净收入=("signed_amt", "sum"),
            转入笔数=("is_in", "sum"),
            转入金额=("in_amt", "sum"),
            转出笔数=("is_out", "sum"),
            转出金额=("out_amt", "sum"),
        )
        .reset_index()
        .rename(columns={"查询对象": "姓名", "交易对方姓名": "对方姓名"})
    )
    total_turnover = summary.groupby("姓名")["交易金额"].transform("sum")
    summary["交易占比%"] = summary["交易金额"] / total_turnover * 100
    summary = summary.sort_values(["姓名", "交易金额"], ascending=[True, False])
    person = summary["姓名"].iat[0]
    out_path = f"2{person}-交易对手分析.xlsx"
    save_df_auto_width(summary, filename=out_path, index=False, engine="openpyxl")
    print(f"✅ 交易对手分析完成，已导出 {out_path}")

    company = summary[summary["对方姓名"].str.contains("公司", na=False)]
    out_path = f"3{person}-与公司相关交易频次分析.xlsx"
    save_df_auto_width(company, filename=out_path, index=False, engine="openpyxl")
    print(f"✅ 与公司相关交易频次分析完成，已导出 {out_path}")
    print("-" * 35  + "end", "-" * 35 + "\n")


def analysis_txn(all_txn: pd.DataFrame) -> None:
    analysis_0 = pd.DataFrame()
    all_txn["交易时间"] = pd.to_datetime(all_txn["交易时间"], errors="coerce")
    all_txn["交易金额"] = pd.to_numeric(all_txn["交易金额"], errors="coerce")
    analysis_0_out = all_txn[all_txn["借贷标志"] == "出"]
    analysis_0_in = all_txn[all_txn["借贷标志"] == "进"]
    amount_range_counts = all_txn["金额区间"].value_counts()
    analysis_0 = pd.DataFrame(
        [
            {
                "交易次数": len(all_txn),
                "交易金额": all_txn["交易金额"].sum(skipna=True),
                "流出额": analysis_0_out["交易金额"].sum(skipna=True),
                "流入额": analysis_0_in["交易金额"].sum(skipna=True),
                "单笔最大支出": analysis_0_out["交易金额"].max(skipna=True),
                "单笔最大收入": analysis_0_in["交易金额"].max(skipna=True),
                "净流入": analysis_0_in["交易金额"].sum(skipna=True)
                - analysis_0_out["交易金额"].sum(skipna=True),
                "最后交易时间": all_txn["交易时间"].max(),
                "0-2千次数": amount_range_counts.get("2000以下", 0),
                "2千-5千次数": amount_range_counts.get("2000-5000", 0),
                "5千-2万次数": amount_range_counts.get("5000-20000", 0),
                "2万-5万次数": amount_range_counts.get("20000-50000", 0),
                "5万以上次数": amount_range_counts.get("50000以上", 0),
            }
        ]
    )
    person = all_txn["查询对象"].iat[0]
    out_path = f"0{person}-资产分析.xlsx"
    save_df_auto_width(analysis_0, filename=out_path, index=False, engine="openpyxl")
    print(f"✅ 资产分析完成，已导出 {out_path}")

    # 存取现1万以上
    analysis_1 = all_txn[
        (
            all_txn["交易类型"].isin(["柜面存款", "柜面取款", "柜面收本金"])
            | all_txn["交易类型"].str.contains("现", na=False)
        )
        & (all_txn["交易金额"] >= 10_000)
    ].drop(columns=["序号", "查询项", "币种"])
    out_path = f"1{person}-存取现1万以上.xlsx"
    save_df_auto_width(analysis_1, filename=out_path, index=False, engine="openpyxl")
    print(f"✅ 存取现1万以上完成，已导出 {out_path}")

    # 大额资金50万以上
    analysis_2 = all_txn[all_txn["交易金额"] >= 500_000].drop(
        columns=["序号", "查询项", "币种"]
    )
    out_path = f"1{person}-大额资金50万以上.xlsx"
    save_df_auto_width(analysis_2, filename=out_path, index=False, engine="openpyxl")
    print(f"✅ 大额资金50万以上完成，已导出 {out_path}")

    # 资金来源分析
    analysis_3 = all_txn.copy()
    analysis_3["is_in"] = analysis_3["借贷标志"] == "进"
    analysis_3["signed_amt"] = analysis_3["交易金额"] * analysis_3["is_in"].map(
        {True: 1, False: -1}
    )
    analysis_3["in_amt"] = analysis_3["交易金额"].where(analysis_3["is_in"], 0)
    analysis_3 = (
        analysis_3.groupby("交易对方姓名", dropna=False)
        .agg(
            交易金额=("交易金额", "sum"),
            交易次数=("交易金额", "size"),
            流入额=("in_amt", "sum"),
            净流入=("signed_amt", "sum"),
            单笔最大收入=("in_amt", "max"),
        )
        .reset_index()
    )
    total_inflow = analysis_3["流入额"].sum()
    analysis_3["流入比%"] = analysis_3["流入额"] / total_inflow * 100
    analysis_3 = analysis_3.sort_values("流入额", ascending=False)
    out_path = f"1{person}-资金来源分析.xlsx"
    save_df_auto_width(analysis_3, filename=out_path, index=False, engine="openpyxl")
    print(f"✅ 资金来源分析完成，已导出 {out_path}")

def create_gui() -> None:
    root = tk.Tk()
    root.title("温岭纪委交易流水分析工具")
    root.minsize(780, 560)

    # ===== 顶部标题与署名 =====
    title_lbl = ttk.Label(
        root, text="温岭纪委交易流水分析工具", font=("仿宋", 20, "bold")
    )
    title_lbl.grid(row=0, column=0, columnspan=3, pady=(15, 0))

    author_lbl = ttk.Label(
        root, text="© 温岭纪委六室 单柳昊", font=("微软雅黑", 9)
    )
    author_lbl.grid(row=1, column=0, columnspan=3, pady=(0, 15))

    # ===== 输入区 =====
    ttk.Label(root, text="工作目录:").grid(
        row=2, column=0, sticky="e", padx=8, pady=8
    )
    path_var = tk.StringVar(value=str(Path.home()))
    ttk.Entry(root, textvariable=path_var, width=60).grid(
        row=2, column=1, padx=5, pady=8, sticky="we"
    )

    def browse_path() -> None:
        p = filedialog.askdirectory(title="选择工作目录")
        if p:
            path_var.set(p)

    ttk.Button(root, text="浏览...", command=browse_path).grid(
        row=2, column=2, padx=5, pady=8
    )

    ttk.Label(root, text="账户姓名:").grid(
        row=3, column=0, sticky="e", padx=8, pady=8
    )
    name_var = tk.StringVar()
    ttk.Entry(root, textvariable=name_var, width=25).grid(
        row=3, column=1, sticky="w", padx=5, pady=8
    )

    # ===== 日志输出区 =====
    log_box = tk.Text(root, width=90, height=18, state="disabled")
    log_box.grid(row=5, column=0, columnspan=3, padx=10, pady=(5, 10), sticky="nsew")

    root.columnconfigure(1, weight=1)
    root.rowconfigure(5, weight=1)

    def log(msg: str):
        log_box.configure(state="normal")
        log_box.insert(
            "end", f"{datetime.datetime.now():%H:%M:%S}  {msg.rstrip()}\n"
        )
        log_box.configure(state="disabled")
        log_box.see("end")
        root.update_idletasks()

    # ===== 运行核心流程 =====
    def run_workflow(data_path: str, holder: str):
        global OUT_DIR
        OUT_DIR = Path(data_path).expanduser().resolve() / f"{holder}-分析"
        OUT_DIR.mkdir(parents=True, exist_ok=True)

        # 将 print 重定向到 log()
        _orig_print = builtins.print

        def _patched_print(*args, **kwargs):
            text = " ".join(str(a) for a in args)
            log(text)

        builtins.print = _patched_print
        try:
            all_txn = merge_txn_by_name(data_path, holder)
            analysis_txn(all_txn)
            make_partner_summary(all_txn)
            messagebox.showinfo("完成", f"分析已完成！\n结果保存在:\n{OUT_DIR}")
        except Exception as exc:
            messagebox.showerror("错误", str(exc))
        finally:
            builtins.print = _orig_print

    # ===== 开始按钮（多线程避免界面卡死）=====
    def on_start():
        path = path_var.get().strip()
        name = name_var.get().strip()
        if not path or not name:
            messagebox.showwarning("提示", "请先选择工作目录并填写账户姓名！")
            return
        threading.Thread(
            target=run_workflow, args=(path, name), daemon=True
        ).start()

    ttk.Button(root, text="开始分析", command=on_start, width=18).grid(
        row=4, column=1, pady=10
    )

    root.mainloop()


# ----------------------------------------------------------------------

if __name__ == "__main__":
    create_gui()
