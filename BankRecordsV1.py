#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
交易流水批量分析工具 GUI   v6-plus
Author  : 温岭纪委六室 单柳昊   （2025-08-05 修订）

"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading, warnings, builtins, datetime, re
from pathlib import Path
from functools import wraps, lru_cache
from typing import Optional, List

import pandas as pd
import numpy as np
from chinese_calendar import is_holiday, is_workday

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ------------------------------------------------------------------
OUT_DIR: Optional[Path] = None
full_ts_pat = re.compile(r"\d{4}-\d{2}-\d{2}-\d{2}\.\d{2}\.\d{2}\.\d+")
TEMPLATE_COLS = [
    "序号","查询对象","反馈单位","查询项","查询账户","查询卡号","交易类型","借贷标志","币种",
    "交易金额","账户余额","交易时间","交易流水号","本方账号","本方卡号","交易对方姓名","交易对方账户",
    "交易对方卡号","交易对方证件号码","交易对手余额","交易对方账号开户行","交易摘要","交易网点名称",
    "交易网点代码","日志号","传票号","凭证种类","凭证号","现金标志","终端号","交易是否成功",
    "交易发生地","商户名称","商户号","IP地址","MAC","交易柜员号","备注",
]

# ------------------------------------------------------------------
# ②   通用工具
# ------------------------------------------------------------------
SKIP_HEADER_KEYWORDS = [
    "反洗钱-电子账户交易明细",
    "信用卡消费明细",
]

def should_skip_special(p: Path) -> str | None:
    """首 3 行包含关键字则返回关键字，否则 None"""
    try:
        head = pd.read_excel(p, header=None, nrows=3)
        for kw in SKIP_HEADER_KEYWORDS:
            if head.astype(str).apply(lambda col: col.str.contains(kw, na=False)).any().any():
                return kw
        return None
    except Exception:
        return None

def _normalize_time(t: str, is_old: bool) -> str:
    if not t:
        return ""
    if "." in t:
        t = re.sub(r"\.(\d{1,6})$", lambda m: ":" + m.group(1).ljust(6, "0"), t.replace(".", ":"))
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
    if OUT_DIR is not None:
        filename = OUT_DIR / filename
    filename = Path(filename).with_suffix(".xlsx")
    filename.parent.mkdir(parents=True, exist_ok=True)

    df = df.replace(np.nan, "")
    if engine == "xlsxwriter":
        with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=index)
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                width = max(
                    min_width,
                    min(max(df[col].astype(str).map(len).max(), len(str(col))) + 2, max_width),
                )
                ws.set_column(i, i, width)
    else:  # openpyxl
        df.to_excel(filename, sheet_name=sheet_name, index=index, engine="openpyxl")
        from openpyxl import load_workbook
        wb = load_workbook(filename)
        ws = wb[sheet_name]
        for col_cells in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells) + 2
            ws.column_dimensions[col_cells[0].column_letter].width = max(
                min_width, min(width, max_width)
            ) + 5
        wb.save(filename)

def str_to_weekday(date_val) -> str:
    dt = pd.to_datetime(date_val, errors="coerce")
    return "wrong" if pd.isna(dt) else ["星期一","星期二","星期三","星期四","星期五","星期六","星期日"][dt.weekday()]

def holiday_status(date_val) -> str:
    dt = pd.to_datetime(date_val, errors="coerce")
    if pd.isna(dt):
        return "wrong"
    d = dt.date()
    try:
        return "节假日" if is_holiday(d) else ("工作日" if is_workday(d) else "周末")
    except Exception:
        return "周末" if dt.weekday() >= 5 else "工作日"

# ------------------------------------------------------------------
# ③   人名识别辅助
# ------------------------------------------------------------------
NAME_CANDIDATE_COLS: List[str] = [
    "账户名称", "户名", "账户名", "账号名称", "账号名", "姓名", "客户名称", "查询对象"
]

def extract_holder_from_df(raw: pd.DataFrame) -> str:
    for col in raw.columns:
        if any(key in col for key in NAME_CANDIDATE_COLS):
            s = raw[col].dropna()
            if not s.empty:
                v = str(s.iloc[0]).strip()
                if v and len(v) <= 10:
                    return v
    return ""

def fallback_holder_from_path(p: Path) -> str:
    name = p.parent.name
    if "农商行" in name:
        name = p.parent.parent.name if p.parent.parent != p.parent else ""
    if not name or "农商行" in name:
        name = re.split(r"[-_]", p.stem)[0]
    return name or "未知"

@lru_cache(maxsize=None)
def holder_from_folder(folder: Path) -> str:
    for fp in folder.glob("*.xls*"):
        try:
            header_idx = _header_row(fp)          # 自动定位表头行
            preview = pd.read_excel(fp, header=header_idx, nrows=5)
            if "账户名称" in preview.columns:
                s = preview["账户名称"].dropna()
                if not s.empty:
                    return str(s.iloc[0]).strip()
        except Exception:
            # 文件打不开或格式异常，直接跳过
            pass
    return ""

# ------------------------------------------------------------------
# ④   解析函数
# ------------------------------------------------------------------
def _header_row(path):
    raw = pd.read_excel(path, header=None, nrows=15)
    for i, r in raw.iterrows():
        if "交易日期" in r.values:
            return i
    return 0

def _read_raw(p: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(p, header=_header_row(p))
    except Exception as e:
        print("❌", p.name, e)
        return pd.DataFrame()

def _parse_dt(d, t, is_old):
    try:
        if full_ts_pat.fullmatch(t.strip()):
            dt = pd.to_datetime(t, format="%Y-%m-%d-%H.%M.%S.%f", errors="coerce")
        else:
            dt = pd.to_datetime(f"{d} {_normalize_time(t, is_old)}".strip(), errors="coerce")
        return dt.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(dt) else "wrong"
    except Exception:
        return "wrong"

# ===============================
# ⑤   泰隆银行 → 模板
# ===============================
def tl_to_template(raw) -> pd.DataFrame:
    """
    泰隆银行流水 → 统一模板字段 TEMPLATE_COLS
    2025-08-06  增强版：
    • 传入  DataFrame  —— 原逻辑不变；返回单 sheet 的模板表
    • 传入 dict[str, DataFrame] —— 自动遍历所有 sheet，
      把各 sheet 解析后纵向合并；并在最前插入 "__sheet__" 列标明来源 sheet
    """

    # ---------- A. 若传入的是 {sheet: DataFrame} ----------
    if isinstance(raw, dict):
        frames = []
        for sheet_name, df_sheet in raw.items():
            one = tl_to_template(df_sheet)          # 递归走单-sheet 逻辑
            if not one.empty:
                one.insert(0, "__sheet__", sheet_name)
                frames.append(one)
        return (
            pd.concat(frames, ignore_index=True)
            if frames
            else pd.DataFrame(columns=TEMPLATE_COLS)
        )

    # ---------- B. 单个 DataFrame：旧逻辑（提炼成内部函数） ----------
    def _one_sheet(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame(columns=TEMPLATE_COLS)

        # ====== 小工具 ======
        def col(c, default=""):
            return df[c] if c in df else pd.Series([default] * len(df), index=df.index)

        def col_multi(keys, default=""):
            for k in keys:
                if k in df:
                    return df[k]
            return pd.Series([default] * len(df), index=df.index)

        out = pd.DataFrame(columns=TEMPLATE_COLS)

        # ====== 基本信息 ======
        out["本方账号"] = out["查询账户"] = col_multi(["客户账号", "账号", "本方账号"], "wrong")
        out["反馈单位"] = "泰隆银行"
        out["查询对象"] = col_multi(["账户名称", "户名", "客户名称"], "wrong")
        out["币种"] = col_multi(["币种", "货币", "币别"]).replace("156", "CNY").fillna("CNY")
        out["借贷标志"] = col_multi(["借贷标志", "借贷方向", "借贷"], "")

        # ====== 金额 / 余额 ======
        debit = pd.to_numeric(col_multi(["借方发生额", "借方发生金额"], 0), errors="coerce")
        credit = pd.to_numeric(col_multi(["贷方发生额", "贷方发生金额"], 0), errors="coerce")
        out["交易金额"] = debit.fillna(0).where(debit.ne(0), credit)
        out["账户余额"] = pd.to_numeric(col_multi(["账户余额", "余额"], 0), errors="coerce")

        # ====== 时间字段 ======
        dates = col_multi(["交易日期", "原交易日期", "会计日期"]).astype(str)
        raw_times = col_multi(["交易时间", "原交易时间", "时间"]).astype(str).str.strip()

        def _tidy_time(s: str) -> str:
            if re.fullmatch(r"0+(\.0+)?", s):
                return ""
            if s.count(".") >= 2:
                p = s.split(".")
                if len(p[0]) == 2 and len(p[1]) == 2 and len(p[2]) == 2:
                    return ".".join(p[:3])
            return s

        def _clean_time(s: str) -> str:
            s = s.strip()
            if re.fullmatch(r"0+(\.0+)?", s):
                return ""
            if re.fullmatch(r"\d{1,9}", s):
                return s.zfill(9)[:6]        # HHMMSS
            return s

        times = raw_times.apply(lambda x: _clean_time(_tidy_time(x)))
        out["交易时间"] = [
            _parse_dt(d, t, is_old=False) for d, t in zip(dates, times)
        ]

        # ====== 其它字段 ======
        out["交易流水号"]        = col_multi(["原柜员流水号", "流水号"])
        out["交易类型"]          = col_multi(["交易码", "交易类型", "业务种类"])
        out["交易对方姓名"]       = col_multi(["对方户名", "交易对手名称"], " ")
        out["交易对方账户"]       = col_multi(["对方客户账号", "对方账号"], " ")
        out["交易对方账号开户行"]   = col_multi(["对方金融机构名称", "对方开户行"], " ")
        out["交易摘要"]          = col_multi(["摘要描述", "摘要"], " ")
        out["交易网点代码"]        = col_multi(["机构号", "网点代码"], " ")
        out["终端号"]           = col_multi(["渠道号", "终端号"], " ")
        out["交易柜员号"]         = col_multi(["柜员号"], " ")
        out["备注"]            = col_multi(["备注", "附言"], " ")

        out["凭证种类"] = col_multi(["凭证类型"], "")
        out["凭证号"]   = col_multi(["凭证序号"], "")

        return out

    # ---------- C. 传入的是 DataFrame ----------
    return _one_sheet(raw)

# ------------------------------------------------------------------
# ⑤   民泰 → 模板
# ------------------------------------------------------------------
def mt_to_template(raw: pd.DataFrame) -> pd.DataFrame:
    """
    民泰银行流水 → 统一模板字段 TEMPLATE_COLS
    2025-08-06 修订：
    1. 自动识别表头行（同时含“时间”“账号卡号”）并去掉前导说明行
    2. 兼容“客户姓名/名称”在说明区的多种写法
    3. 新增：过滤包含 “支出笔数 / 收入笔数” 等文字的汇总行
    """
    if raw.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    # ---------- ① 识别表头行 ----------
    header_idx = None
    for i, row in raw.iterrows():
        cells = row.astype(str).str.strip().tolist()
        if "时间" in cells and "账号卡号" in cells:
            header_idx = i
            break
    if header_idx is None:
        for i, row in raw.iterrows():
            if row.astype(str).str.contains("序号").any():
                header_idx = i
                break
    if header_idx is None:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    # ---------- ② 提取客户姓名 ----------
    holder = ""
    name_inline = re.compile(r"客户(?:姓名|名称)\s*[:：]?\s*([^\s:：]{2,10})")
    for i in range(header_idx):
        vals = raw.iloc[i].astype(str).tolist()
        for j, cell in enumerate(vals):
            cs = cell.strip()
            m = name_inline.match(cs)
            if m:                                    # 同格：A1 = “客户名称:张三”
                holder = m.group(1)
                break
            if re.fullmatch(r"客户(?:姓名|名称)\s*[:：]?", cs):  # 分列：A1 = “客户名称:”
                nxt = str(vals[j + 1]).strip() if j + 1 < len(vals) else ""
                if nxt and nxt.lower() != "nan":
                    holder = nxt
                    break
        if holder:
            break
    holder = holder or "未知"

    # ---------- ③ 数据区 ----------
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = raw.iloc[header_idx].astype(str).str.strip()
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # ★ 过滤“支出笔数 / 收入笔数 …”汇总行
    summary_mask = df.apply(
        lambda row: row.astype(str)
        .str.contains(r"支出笔数|收入笔数|支出累计金额|收入累计金额")
        .any(),
        axis=1,
    )
    df = df[~summary_mask].copy()

    def col(c, default=""):
        return df[c] if c in df else pd.Series(default, index=df.index)

    out = pd.DataFrame(columns=TEMPLATE_COLS)

    # ===== 基本信息 =====
    acct = col("账号卡号").astype(str).str.replace(r"\.0$", "", regex=True)
    out["本方账号"] = out["查询账户"] = acct
    out["查询对象"] = holder
    out["反馈单位"] = "民泰银行"
    out["币种"] = col("币种").astype(str).replace("人民币", "CNY").fillna("CNY")

    # ===== 金额 / 借贷标志 =====
    debit  = pd.to_numeric(col("支出"), errors="coerce").fillna(0)
    credit = pd.to_numeric(col("收入"), errors="coerce").fillna(0)
    out["交易金额"] = credit.where(credit.gt(0), -debit)
    out["账户余额"] = pd.to_numeric(col("余额"), errors="coerce")
    out["借贷标志"] = np.where(credit.gt(0), "进", "出")

    # ===== 交易时间 =====
    def _fmt_time(v: str) -> str:
        v = str(v).strip()
        try:
            return datetime.datetime.strptime(v, "%Y%m%d %H:%M:%S").strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        except Exception:
            return v or "wrong"

    out["交易时间"] = col("时间").astype(str).apply(_fmt_time)

    # ===== 其它字段 =====
    out["交易摘要"]        = col("摘要", " ")
    out["交易流水号"]      = col("柜员流水号").astype(str).str.strip()
    out["交易柜员号"]       = col("记账柜员 ").astype(str).str.strip()
    out["交易网点代码"]      = col("记账机构").astype(str).str.strip()
    out["交易对方姓名"]     = col("交易对手名称", " ").astype(str).str.strip()
    out["交易对方账户"]     = col("交易对手账号", " ").astype(str).str.strip()
    out["交易对方账号开户行"] = col("交易对手行名", " ").astype(str).str.strip()
    out["终端号"]         = col("交易渠道")
    out["备注"]          = col("附言", " ")

    return out
# ------------------------------------------------------------------
# ⑤   农商行 → 模板
# ------------------------------------------------------------------
def rc_to_template(raw: pd.DataFrame, holder: str, is_old: bool) -> pd.DataFrame:
    if raw.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)
    def col(c, default=""):
        return raw[c] if c in raw else pd.Series([default] * len(raw), index=raw.index)

    out = pd.DataFrame(columns=TEMPLATE_COLS)
    out["本方账号"] = out["查询账户"] = col("账号", "wrong")
    out["交易金额"] = col("发生额") if is_old else col("交易金额")
    out["账户余额"] = col("余额") if is_old else col("交易余额")
    out["反馈单位"] = "老农商银行" if is_old else "新农商银行"

    dates = col("交易日期").astype(str)
    times = col("交易时间").astype(str)
    out["交易时间"] = [_parse_dt(d, t, is_old) for d, t in zip(dates, times)]

    out["借贷标志"] = col("借贷标志")          # 保留原值
    out["币种"] = "CNY" if is_old else col("币种").replace("人民币", "CNY")
    out["查询对象"] = holder
    out["交易对方姓名"] = col("对方姓名", " ")
    out["交易对方账户"] = col("对方账号", " ")
    out["交易网点名称"] = col("代理行机构号") if is_old else col("交易机构")
    out["交易摘要"] = col("备注") if is_old else col("摘要", "wrong")
    return out

# ------------------------------------------------------------------
# ⑥   合并全部流水
# ------------------------------------------------------------------
def merge_all_txn(root_dir: str) -> pd.DataFrame:
    root = Path(root_dir).expanduser().resolve()

    # 网上银行：命名规则 "*-*-交易流水.xls*"
    china_files = [p for p in root.rglob("*-*-交易流水.xls*")]

    # 农商行文件
    all_excel = list(root.rglob("*.xls*"))
    rc_candidates = [p for p in all_excel if "农商行" in p.as_posix()]
    pattern_old = re.compile(r"老\s*[账帐]\s*(?:号|户)")
    old_rc = [p for p in rc_candidates if pattern_old.search(p.stem)]
    new_rc = [p for p in rc_candidates if p not in old_rc]

    tl_files = [p for p in all_excel if "泰隆" in p.as_posix()]
    mt_files = [p for p in all_excel if "民泰" in p.as_posix()] 
    # tz_files = [p for p in all_excel if "台州银行" in p.as_posix()]

    print(
        f"✅ 网上银行 {len(china_files)} 份，"
        f"老农商 {len(old_rc)} 份，新农商 {len(new_rc)} 份，"
        f"泰隆银行 {len(tl_files)} 份，"
        f"民泰银行 {len(mt_files)} 份"
        # f"台州银行 {len(tz_files)} 份 "
    )

    dfs = []

    # -------------------- 网上银行 --------------------
    for p in china_files:
        print(f"正在处理 {p.name} ...")
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
            print("❌", p.name, e)

    # -------------------- 农商行 --------------------
    for p in old_rc + new_rc:
        print(f"正在处理 {p.name} ...")
        kw = should_skip_special(p)
        if kw:
            print(f"⏩ 跳过【{p.name}】：表头含“{kw}”")
            continue

        raw = _read_raw(p)

        # ① 本文件
        holder = extract_holder_from_df(raw)
        # ② 同文件夹缓存
        if not holder:
            holder = holder_from_folder(p.parent)
        # ③ fallback
        if not holder:
            holder = fallback_holder_from_path(p)

        is_old = p in old_rc
        df = rc_to_template(raw, holder, is_old)
        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)

    # ----------- 泰隆银行 -----------

    for p in tl_files:
        if "开户" in p.stem:
            continue
        print(f"正在处理 {p.name} ...")

        xls = pd.ExcelFile(p)
        xls_dict = {}
        for sht in xls.sheet_names:
            try:
                header_idx = _header_row(p)
                df_sheet   = pd.read_excel(p, sheet_name=sht, header=header_idx)
                xls_dict[sht] = df_sheet
            except Exception as e:
                print("❌", f"{p.name} -> {sht}", e)

        df = tl_to_template(xls_dict)

        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)
    
    # ----------- 民泰银行 -----------
    for p in mt_files:
        print(f"正在处理 {p.name} ...")
        raw = _read_raw(p)
        df  = mt_to_template(raw)
        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)

    # # ----------- 台州银行 -----------
    # for p in tz_files:
    #     print(f"正在处理 {p.name} ...")
    #     raw = _read_raw(p)
    #     df  = tz_to_template(raw)
    #     if not df.empty:
    #         df["来源文件"] = p.name
    #         dfs.append(df)
    
    print(f"文件读取已完成，正在整合分析！ ...")

    all_txn = pd.concat(dfs, ignore_index=True)

    # ---- 排序、补充列 ----
    ts = pd.to_datetime(all_txn["交易时间"], errors="coerce")
    all_txn.insert(0, "__ts__", ts)
    all_txn.sort_values("__ts__", inplace=True)
    all_txn["序号"] = range(1, len(all_txn) + 1)
    all_txn.drop(columns="__ts__", inplace=True)

    # ⚠️ 统一金额类型
    all_txn["交易金额"] = pd.to_numeric(all_txn["交易金额"], errors="coerce")

    all_txn["星期"] = all_txn["交易时间"].apply(str_to_weekday)
    all_txn["节假日"] = all_txn["交易时间"].apply(holiday_status)

    # ---- 借贷标志最终统一 ----
    def _std_flag(x):
        if pd.isna(x):
            return x
        s = str(x).strip()
        if s in {"1", "借", "D"}:
            return "出"
        if s in {"2", "贷", "C"}:
            return "进"
        return s
    all_txn["借贷标志"] = all_txn["借贷标志"].apply(_std_flag)

    # ---- 金额区间 ----
    bins = [-np.inf, 2000, 5000, 20000, 50000, np.inf]
    labels = ["2000以下", "2000-5000", "5000-20000", "20000-50000", "50000以上"]
    all_txn["金额区间"] = pd.cut(
        all_txn["交易金额"], bins=bins, labels=labels, right=False, include_lowest=True
    )

    save_df_auto_width(all_txn, "所有人-合并交易流水", index=False, engine="openpyxl")
    print("✅ 已导出 所有人-合并交易流水.xlsx")
    return all_txn

# ------------------------------------------------------------------
# ⑦   单人资产 / 对手分析 （略，保持不变）
# ------------------------------------------------------------------
def analysis_txn(df: pd.DataFrame) -> None:
    if df.empty:
        return
    df = df.copy()
    df["交易时间"] = pd.to_datetime(df["交易时间"], errors="coerce")
    df["交易金额"] = pd.to_numeric(df["交易金额"], errors="coerce")
    person = df["查询对象"].iat[0] or "未知"
    prefix = f"{person}/"

    out_df = df[df["借贷标志"] == "出"]
    in_df = df[df["借贷标志"] == "进"]
    counts = df["金额区间"].value_counts()

    summary = pd.DataFrame(
        [
            {
                "交易次数": len(df),
                "交易金额": df["交易金额"].sum(skipna=True),
                "流出额": out_df["交易金额"].sum(skipna=True),
                "流入额": in_df["交易金额"].sum(skipna=True),
                "单笔最大支出": out_df["交易金额"].max(skipna=True),
                "单笔最大收入": in_df["交易金额"].max(skipna=True),
                "净流入": in_df["交易金额"].sum(skipna=True)
                - out_df["交易金额"].sum(skipna=True),
                "最后交易时间": df["交易时间"].max(),
                "0-2千次数": counts.get("2000以下", 0),
                "2千-5千次数": counts.get("2000-5000", 0),
                "5千-2万次数": counts.get("5000-20000", 0),
                "2万-5万次数": counts.get("20000-50000", 0),
                "5万以上次数": counts.get("50000以上", 0),
            }
        ]
    )
    save_df_auto_width(summary, f"{prefix}0{person}-资产分析", index=False, engine="openpyxl")

    cash = df[
        (df["交易类型"].fillna("").str.contains("柜面|现"))
        & (pd.to_numeric(df["交易金额"], errors="coerce") >= 10_000)
    ]
    save_df_auto_width(cash, f"{prefix}1{person}-存取现1万以上", index=False, engine="openpyxl")

    big = df[pd.to_numeric(df["交易金额"], errors="coerce") >= 500_000]
    save_df_auto_width(big, f"{prefix}1{person}-大额资金50万以上", index=False, engine="openpyxl")

    src = df.copy()
    src["is_in"] = src["借贷标志"] == "进"
    src["signed_amt"] = (
        pd.to_numeric(src["交易金额"], errors="coerce")
        * src["is_in"].map({True: 1, False: -1})
    )
    src["in_amt"] = pd.to_numeric(src["交易金额"], errors="coerce").where(src["is_in"], 0)
    src = (
        src.groupby("交易对方姓名", dropna=False)
        .agg(
            交易金额=("交易金额", "sum"),
            交易次数=("交易金额", "size"),
            流入额=("in_amt", "sum"),
            净流入=("signed_amt", "sum"),
            单笔最大收入=("in_amt", "max"),
        )
        .reset_index()
    )
    total = src["流入额"].sum()
    src["流入比%"] = src["流入额"] / total * 100
    src.sort_values("流入额", ascending=False, inplace=True)
    save_df_auto_width(src, f"{prefix}1{person}-资金来源分析", index=False, engine="openpyxl")

def make_partner_summary(df: pd.DataFrame) -> None:
    if df.empty:
        return
    person = df["查询对象"].iat[0] or "未知"
    prefix = f"{person}/"
    d = df.copy()
    d["交易金额"] = pd.to_numeric(d["交易金额"], errors="coerce").fillna(0)
    d["is_in"] = d["借贷标志"] == "进"
    d["abs_amt"] = d["交易金额"].abs()
    d["signed_amt"] = d["交易金额"] * d["is_in"].map({True: 1, False: -1})
    d["in_amt"] = d["交易金额"].where(d["is_in"], 0)
    d["out_amt"] = d["交易金额"].where(~d["is_in"], 0)
    d["gt10k"] = (d["abs_amt"] >= 10_000).astype(int)
    summ = (
        d.groupby(["查询对象", "交易对方姓名"], dropna=False)
        .agg(
            交易次数=("交易金额", "size"),
            交易金额=("abs_amt", "sum"),
            万元以上交易次数=("gt10k", "sum"),
            净收入=("signed_amt", "sum"),
            转入笔数=("is_in", "sum"),
            转入金额=("in_amt", "sum"),
            转出笔数=("is_in", lambda x: (~x).sum()),
            转出金额=("out_amt", "sum"),
        )
        .reset_index()
        .rename(columns={"查询对象": "姓名", "交易对方姓名": "对方姓名"})
    )
    total = summ.groupby("姓名")["交易金额"].transform("sum")
    summ["交易占比%"] = summ["交易金额"] / total * 100
    summ.sort_values(["姓名", "交易金额"], ascending=[True, False], inplace=True)
    save_df_auto_width(summ, f"{prefix}2{person}-交易对手分析", index=False, engine="openpyxl")
    comp = summ[summ["对方姓名"].str.contains("公司", na=False)]
    save_df_auto_width(comp, f"{prefix}3{person}-与公司相关交易频次分析", index=False, engine="openpyxl")

# ------------------------------------------------------------------
# ⑧   GUI
# ------------------------------------------------------------------
def create_gui():
    root = tk.Tk()
    root.title("温岭纪委交易流水批量分析工具")
    root.minsize(780, 560)
    ttk.Label(root, text="温岭纪委交易流水批量分析工具", font=("仿宋", 20, "bold")).grid(
        row=0, column=0, columnspan=3, pady=(15, 0)
    )
    ttk.Label(root, text="© 温岭纪委六室 单柳昊", font=("微软雅黑", 9)).grid(
        row=1, column=0, columnspan=3, pady=(0, 15)
    )

    ttk.Label(root, text="工作目录:").grid(row=2, column=0, sticky="e", padx=8, pady=8)
    path_var = tk.StringVar(value=str(Path.home()))
    ttk.Entry(root, textvariable=path_var, width=60).grid(
        row=2, column=1, sticky="we", padx=5, pady=8
    )
    ttk.Button(
        root,
        text="浏览...",
        command=lambda: path_var.set(
            filedialog.askdirectory(title="选择工作目录") or path_var.get()
        ),
    ).grid(row=2, column=2, padx=5, pady=8)

    log_box = tk.Text(root, width=90, height=18, state="disabled")
    log_box.grid(row=4, column=0, columnspan=3, padx=10, pady=(5, 10), sticky="nsew")
    root.columnconfigure(1, weight=1)
    root.rowconfigure(4, weight=1)

    def log(msg):
        log_box.config(state="normal")
        log_box.insert("end", f"{datetime.datetime.now():%H:%M:%S}  {msg}\n")
        log_box.config(state="disabled")
        log_box.see("end")

    def run(path):
        global OUT_DIR
        OUT_DIR = Path(path).expanduser().resolve() / "批量分析结果"
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        _orig_print = builtins.print
        builtins.print = lambda *a, **k: log(" ".join(map(str, a)))
        try:
            all_txn = merge_all_txn(path)
            if all_txn.empty:
                messagebox.showinfo("完成", "未找到可分析文件")
                return
            for person, df_person in all_txn.groupby("查询对象", dropna=False):
                print(f"--- 分析 {person} ---")
                analysis_txn(df_person)
                make_partner_summary(df_person)
            messagebox.showinfo("完成", f"全部分析完成！结果在:\n{OUT_DIR}")
        except Exception as e:
            messagebox.showerror("错误", str(e))
        finally:
            builtins.print = _orig_print

    def on_start():
        p = path_var.get().strip()
        if not p:
            messagebox.showwarning("提示", "请先选择工作目录！")
            return
        threading.Thread(target=run, args=(p,), daemon=True).start()

    ttk.Button(root, text="开始分析", command=on_start, width=18).grid(row=3, column=1, pady=10)
    root.mainloop()

# ------------------------------------------------------------------
if __name__ == "__main__":
    create_gui()
