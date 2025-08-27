#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
交易流水批量分析工具 GUI   v6-plus
Author  : 温岭纪委六室 单柳昊   （2025-08-05 修订）

（2025-08-27 增补）
- 新增：支持读取同目录下固定文件名 “交易明细信息.csv”
- “查询对象”自动来自同目录 “人员信息.csv” 的“客户姓名”
- “反馈单位”来自 CSV 文件父文件夹名
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
            if head.astype(str).apply(lambda col: col.astype(str).str.contains(kw, na=False)).any().any():
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

# ----------（新增）CSV/人员信息相关通用函数 ----------
def _read_csv_smart(p: Path, **kwargs) -> pd.DataFrame:
    """智能编码读取 CSV：优先 utf-8-sig，其次 gb18030，再退回 utf-8/cp936。"""
    enc_try = ["utf-8-sig", "gb18030", "utf-8", "cp936"]
    last_err = None
    for enc in enc_try:
        try:
            return pd.read_csv(p, encoding=enc, **kwargs)
        except Exception as e:
            last_err = e
    raise last_err or RuntimeError(f"无法读取CSV: {p}")

def _person_from_people_csv(dirpath: Path) -> str:
    """同目录 ‘人员信息.csv’ 中优先取列 ‘客户姓名’ 的首个非空值；提供稳健兜底。"""
    people = dirpath / "人员信息.csv"
    if not people.exists():
        return ""
    try:
        df = _read_csv_smart(people)
    except Exception:
        return ""
    # 直接列名命中
    for col in ["客户姓名", "姓名", "客户名称", "户名"]:
        if col in df.columns:
            ser = df[col].astype(str).str.strip()
            ser = ser[(ser != "") & (ser.str.lower() != "nan")]
            if not ser.empty:
                return ser.iloc[0][:10]
    # 表格里混写的“客户姓名:张三”
    name_pat = re.compile(r"客户(?:姓名|名称)\s*[:：]?\s*([^\s:：]{2,10})")
    for val in df.astype(str).replace("nan", "", regex=False).to_numpy().ravel().tolist():
        m = name_pat.search(val.strip())
        if m:
            return m.group(1)
    return ""

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
        if isinstance(t, str) and full_ts_pat.fullmatch(t.strip()):
            dt = pd.to_datetime(t, format="%Y-%m-%d-%H.%M.%S.%f", errors="coerce")
        else:
            dt = pd.to_datetime(f"{d} {_normalize_time(str(t), is_old)}".strip(), errors="coerce")
        return dt.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(dt) else "wrong"
    except Exception:
        return "wrong"

# ----------（修复&增强）交易明细 CSV → 模板 ----------
def csv_to_template(raw: pd.DataFrame, holder: str, feedback_unit: str) -> pd.DataFrame:
    """
    将‘交易明细信息.csv’映射成统一模板；字段尽量对齐，不强依赖固定表头。
    - 查询对象：传入 holder
    - 反馈单位：传入 feedback_unit（父目录名）
    - 若任一字段转换过程中出错：该字段整列填入 'wrong' 并继续
    """
    if raw.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    # 整体兜底：函数级异常直接返回全 'wrong'
    try:
        df = raw.copy()
        df.columns = pd.Index(df.columns).astype(str).str.strip()
        n = len(df)

        def _S(default=""):
            return pd.Series([default] * n, index=df.index)

        def _safe(name, fn):
            """单列安全赋值：异常→整列'wrong'"""
            try:
                return fn()
            except Exception as e:
                print(f"⚠️ CSV字段[{name}]解析异常：{e}")
                return _S("wrong")

        def col(keys, default=""):
            """支持传入单列名或候选列名列表；不做 .str 操作"""
            if isinstance(keys, str):
                return df[keys] if keys in df else _S(default)
            for k in keys:
                if k in df:
                    return df[k]
            return _S(default)

        def _to_str_no_sci(x):
            """账号/卡号等数字安全转字符串（防科学计数/去'.0'）"""
            if pd.isna(x):
                return ""
            s = str(x).strip()
            if s.lower() == "nan":
                return ""
            if re.fullmatch(r"\d+\.0", s):
                return s[:-2]
            try:
                if isinstance(x, (int, np.integer)):
                    return str(int(x))
                if isinstance(x, float):
                    return f"{x:.0f}" if np.isfinite(x) else ""
                if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
                    return f"{float(s):.0f}"
            except Exception:
                pass
            return s

        def _std_success(v):
            s = str(v).strip()
            if s in {"1","Y","y","是","成功","True","true"}: return "成功"
            if s in {"0","N","n","否","失败","False","false"}: return "失败"
            return "" if s.lower() == "nan" else s

        out = pd.DataFrame(columns=TEMPLATE_COLS, index=df.index)

        # ===== 本方账号/卡号 + 查询账户/卡号 =====
        out["本方账号"] = _safe("本方账号", lambda: col(["交易账号","查询账户","本方账号","账号","账号/卡号","账号卡号"]).map(_to_str_no_sci))
        out["本方卡号"] = _safe("本方卡号", lambda: col(["交易卡号","查询卡号","本方卡号","卡号"]).map(_to_str_no_sci))
        out["查询账户"] = out["本方账号"]
        out["查询卡号"] = out["本方卡号"]

        # ===== 对方账号/卡号（按 类型 列分流）=====
        opp_no  = _safe("交易对手账卡号", lambda: col(["交易对手账卡号","交易对手账号","对方账号","对方账户"]).map(_to_str_no_sci))
        opp_typ = col(["交易对方帐卡号类型","账号/卡号类型"], "")
        typ_s   = opp_typ.astype(str)
        is_card = _safe("交易对方帐卡号类型", lambda: typ_s.str.contains("卡", na=False) | typ_s.isin(["2","卡","卡号"]))
        out["交易对方卡号"] = np.where(is_card, opp_no, "")
        out["交易对方账户"] = np.where(is_card, "", opp_no)

        # ===== 基本信息 =====
        out["查询对象"] = holder or "未知"
        out["反馈单位"] = feedback_unit or "未知"
        out["币种"] = _safe("币种", lambda: col(["交易币种","币种","币别","货币"], "CNY").astype(str).replace(
            {"人民币":"CNY","RMB":"CNY","156":"CNY"}).fillna("CNY"))

        # ===== 金额 / 余额 =====
        amt = _safe("交易金额", lambda: pd.to_numeric(col(["交易金额","金额","发生额"], 0), errors="coerce"))
        out["交易金额"] = amt
        out["账户余额"] = _safe("账户余额", lambda: pd.to_numeric(col(["交易余额","余额","账户余额"], 0), errors="coerce"))

        # ===== 借贷/收付 → 进/出 =====
        jl = col(["收付标志","借贷标志","借贷方向","借贷"], "")
        out["借贷标志"] = col(["收付标志",""])

        # ===== 交易时间 =====
        if "交易时间" in df.columns:
            out["交易时间"] = _safe("交易时间", lambda: np.where(
                pd.to_datetime(df["交易时间"], errors="coerce").notna(),
                pd.to_datetime(df["交易时间"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M:%S"),
                df["交易时间"].astype(str)
            ))
        else:
            out["交易时间"] = _S("wrong")  # 没有交易时间列则记为 wrong

        # ===== 其它字段对齐 =====
        out["交易类型"]              = col(["交易类型","业务种类","交易码"], "")
        out["交易流水号"]            = col(["交易流水号","柜员流水号","流水号"], "")
        out["交易对方姓名"]           = col(["对手户名","交易对手名称","对手方名称","对方户名","对方名称","对方姓名","收/付方名称"], " ")
        out["交易对方证件号码"]         = col(["对手身份证号","对方证件号码"], " ")
        out["交易对手余额"]           = _safe("交易对手余额", lambda: pd.to_numeric(col(["对手交易余额"], ""), errors="coerce"))
        out["交易对方账号开户行"]        = col(["对手开户银行","交易对手行名","对方开户行","对方金融机构名称"], " ")
        out["交易摘要"]              = col(["摘要说明","交易摘要","摘要","附言","用途"], " ")
        out["交易网点名称"]            = col(["交易网点名称","交易机构","网点名称"], "")
        out["交易网点代码"]            = col(["交易网点代码","机构号","网点代码"], "")
        out["日志号"]               = col(["日志号"], "")
        out["传票号"]               = col(["传票号"], "")
        out["凭证种类"]              = col(["凭证种类","凭证类型"], "")
        out["凭证号"]               = col(["凭证号","凭证序号"], "")
        out["现金标志"]              = col(["现金标志"], "")
        out["终端号"]               = col(["终端号","渠道号"], "")
        succ                        = col(["交易是否成功","查询反馈结果"], "")
        out["交易是否成功"]            = succ.map(_std_success)
        out["交易发生地"]             = col(["交易发生地","交易场所"], "")
        out["商户名称"]              = col(["商户名称"], "")
        out["商户号"]               = col(["商户号"], "")
        out["IP地址"]              = col(["IP地址"], "")
        out["MAC"]                = col(["MAC地址","MAC"], "")
        out["交易柜员号"]             = col(["交易柜员号","柜员号","记账柜员"], "")

        # ===== 备注：并入“查询反馈结果原因”；异常→wrong =====
        try:
            beizhu = col(["备注","附言","说明"], "").astype(str)
            reason = col(["查询反馈结果原因"], "").astype(str)
            beizhu_clean = beizhu.where(~beizhu.str.lower().eq("nan"), "")
            reason_clean = reason.where(~reason.str.lower().eq("nan"), "")
            out["备注"] = np.where(
                reason_clean != "",
                np.where(beizhu_clean != "", beizhu_clean + "｜原因：" + reason_clean, "原因：" + reason_clean),
                beizhu_clean,
            )
        except Exception as e:
            print(f"⚠️ CSV字段[备注/原因]解析异常：{e}")
            out["备注"] = _S("wrong")

        return out

    except Exception as e:
        # 整体兜底：构造一个全 'wrong' 的模板（保持原有行数）
        print(f"❌ CSV转模板发生异常：{e}")
        n = len(raw)
        bad = pd.DataFrame({col: ["wrong"] * n for col in TEMPLATE_COLS})
        return bad

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
    if isinstance(raw, dict):
        frames = []
        for sheet_name, df_sheet in raw.items():
            one = tl_to_template(df_sheet)
            if not one.empty:
                one.insert(0, "__sheet__", sheet_name)
                frames.append(one)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=TEMPLATE_COLS)

    def _one_sheet(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame(columns=TEMPLATE_COLS)

        def col(c, default=""):
            return df[c] if c in df else pd.Series([default] * len(df), index=df.index)

        def col_multi(keys, default=""):
            for k in keys:
                if k in df:
                    return df[k]
            return pd.Series([default] * len(df), index=df.index)

        out = pd.DataFrame(columns=TEMPLATE_COLS)
        out["本方账号"] = out["查询账户"] = col_multi(["客户账号","账号","本方账号"], "wrong")
        out["反馈单位"] = "泰隆银行"
        out["查询对象"] = col_multi(["账户名称","户名","客户名称"], "wrong")
        out["币种"] = col_multi(["币种","货币","币别"]).replace("156","CNY").fillna("CNY")
        out["借贷标志"] = col_multi(["借贷标志","借贷方向","借贷"], "")

        debit  = pd.to_numeric(col_multi(["借方发生额","借方发生金额"], 0), errors="coerce")
        credit = pd.to_numeric(col_multi(["贷方发生额","贷方发生金额"], 0), errors="coerce")
        out["交易金额"] = debit.fillna(0).where(debit.ne(0), credit)
        out["账户余额"] = pd.to_numeric(col_multi(["账户余额","余额"], 0), errors="coerce")

        dates = col_multi(["交易日期","原交易日期","会计日期"]).astype(str)
        raw_times = col_multi(["交易时间","原交易时间","时间"]).astype(str).str.strip()

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
                return s.zfill(9)[:6]
            return s

        times = raw_times.apply(lambda x: _clean_time(_tidy_time(x)))
        out["交易时间"] = [_parse_dt(d, t, is_old=False) for d, t in zip(dates, times)]

        out["交易流水号"]        = col_multi(["原柜员流水号","流水号"])
        out["交易类型"]          = col_multi(["交易码","交易类型","业务种类"])
        out["交易对方姓名"]       = col_multi(["对方户名","交易对手名称"], " ")
        out["交易对方账户"]       = col_multi(["对方客户账号","对方账号"], " ")
        out["交易对方账号开户行"]   = col_multi(["对方金融机构名称","对方开户行"], " ")
        out["交易摘要"]          = col_multi(["摘要描述","摘要"], " ")
        out["交易网点代码"]        = col_multi(["机构号","网点代码"], " ")
        out["终端号"]           = col_multi(["渠道号","终端号"], " ")
        out["交易柜员号"]         = col_multi(["柜员号"], " ")
        out["备注"]            = col_multi(["备注","附言"], " ")

        out["凭证种类"] = col_multi(["凭证类型"], "")
        out["凭证号"]   = col_multi(["凭证序号"], "")
        return out

    return _one_sheet(raw)

# ------------------------------------------------------------------
# ⑤   民泰 → 模板
# ------------------------------------------------------------------
def mt_to_template(raw: pd.DataFrame) -> pd.DataFrame:
    if raw.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)

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

    holder = ""
    name_inline = re.compile(r"客户(?:姓名|名称)\s*[:：]?\s*([^\s:：]{2,10})")
    for i in range(header_idx):
        vals = raw.iloc[i].astype(str).tolist()
        for j, cell in enumerate(vals):
            cs = cell.strip()
            m = name_inline.match(cs)
            if m:
                holder = m.group(1)
                break
            if re.fullmatch(r"客户(?:姓名|名称)\s*[:：]?", cs):
                nxt = str(vals[j+1]).strip() if j+1 < len(vals) else ""
                if nxt and nxt.lower() != "nan":
                    holder = nxt
                    break
        if holder:
            break
    holder = holder or "未知"

    df = raw.iloc[header_idx + 1:].copy()
    df.columns = raw.iloc[header_idx].astype(str).str.strip()
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    summary_mask = df.apply(
        lambda row: row.astype(str).str.contains(r"支出笔数|收入笔数|支出累计金额|收入累计金额").any(),
        axis=1,
    )
    df = df[~summary_mask].copy()

    def col(c, default=""):
        return df[c] if c in df else pd.Series(default, index=df.index)

    out = pd.DataFrame(columns=TEMPLATE_COLS)
    acct = col("账号卡号").astype(str).str.replace(r"\.0$", "", regex=True)
    out["本方账号"] = out["查询账户"] = acct
    out["查询对象"] = holder
    out["反馈单位"] = "民泰银行"
    out["币种"] = col("币种").astype(str).replace("人民币","CNY").fillna("CNY")

    debit  = pd.to_numeric(col("支出"), errors="coerce").fillna(0)
    credit = pd.to_numeric(col("收入"), errors="coerce").fillna(0)
    out["交易金额"] = credit.where(credit.gt(0), -debit)
    out["账户余额"] = pd.to_numeric(col("余额"), errors="coerce")
    out["借贷标志"] = np.where(credit.gt(0), "进", "出")

    def _fmt_time(v: str) -> str:
        v = str(v).strip()
        try:
            return datetime.datetime.strptime(v, "%Y%m%d %H:%M:%S").strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return v or "wrong"

    out["交易时间"] = col("时间").astype(str).apply(_fmt_time)

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

    out["借贷标志"] = col("借贷标志")
    out["币种"] = "CNY" if is_old else col("币种").replace("人民币","CNY")
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

    china_files = [p for p in root.rglob("*-*-交易流水.xls*")]

    all_excel = list(root.rglob("*.xls*"))
    rc_candidates = [p for p in all_excel if "农商行" in p.as_posix()]
    pattern_old = re.compile(r"老\s*[账帐]\s*(?:号|户)")
    old_rc = [p for p in rc_candidates if pattern_old.search(p.stem)]
    new_rc = [p for p in rc_candidates if p not in old_rc]

    tl_files = [p for p in all_excel if "泰隆" in p.as_posix()]
    mt_files = [p for p in all_excel if "民泰" in p.as_posix()]

    csv_txn_files = [p for p in root.rglob("交易明细信息.csv")]

    print(
        f"✅ 网上银行 {len(china_files)} 份，"
        f"老农商 {len(old_rc)} 份，新农商 {len(new_rc)} 份，"
        f"泰隆银行 {len(tl_files)} 份，"
        f"民泰银行 {len(mt_files)} 份，"
        f"交易明细CSV {len(csv_txn_files)} 份"
    )

    dfs = []

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

    for p in old_rc + new_rc:
        print(f"正在处理 {p.name} ...")
        kw = should_skip_special(p)
        if kw:
            print(f"⏩ 跳过【{p.name}】：表头含“{kw}”")
            continue

        raw = _read_raw(p)

        holder = extract_holder_from_df(raw)
        if not holder:
            holder = holder_from_folder(p.parent)
        if not holder:
            holder = fallback_holder_from_path(p)

        is_old = p in old_rc
        df = rc_to_template(raw, holder, is_old)
        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)

    for p in tl_files:
        if "开户" in p.stem:
            continue
        print(f"正在处理 {p.name} ...")
        try:
            xls = pd.ExcelFile(p)
        except Exception as e:
            print("❌", f"{p.name} 载入失败", e)
            continue

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

    for p in mt_files:
        print(f"正在处理 {p.name} ...")
        raw = _read_raw(p)
        df  = mt_to_template(raw)
        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)

    for p in csv_txn_files:
        print(f"正在处理 {p.name} ...")
        try:
            raw_csv = _read_csv_smart(p)
        except Exception as e:
            print("❌ 无法读取CSV", p.name, e)
            continue

        holder = _person_from_people_csv(p.parent)
        if not holder:
            holder = holder_from_folder(p.parent) or fallback_holder_from_path(p)

        feedback_unit = p.parent.name
        try:
            df = csv_to_template(raw_csv, holder, feedback_unit)
        except Exception as e:
            print("❌ CSV转模板失败", p.name, e)
            continue

        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)

    print(f"文件读取已完成，正在整合分析！ ...")

    if not dfs:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    all_txn = pd.concat(dfs, ignore_index=True)

    ts = pd.to_datetime(all_txn["交易时间"], errors="coerce")
    all_txn.insert(0, "__ts__", ts)
    all_txn.sort_values("__ts__", inplace=True)
    all_txn["序号"] = range(1, len(all_txn) + 1)
    all_txn.drop(columns="__ts__", inplace=True)

    all_txn["交易金额"] = pd.to_numeric(all_txn["交易金额"], errors="coerce")

    all_txn["星期"] = all_txn["交易时间"].apply(str_to_weekday)
    all_txn["节假日"] = all_txn["交易时间"].apply(holiday_status)

    def _std_flag(x):
        if pd.isna(x):
            return x
        s = str(x).strip()
        if s in {"1","借","D"}: return "出"
        if s in {"2","贷","C"}: return "进"
        return s
    all_txn["借贷标志"] = all_txn["借贷标志"].apply(_std_flag)

    bins = [-np.inf, 2000, 5000, 20000, 50000, np.inf]
    labels = ["2000以下","2000-5000","5000-20000","20000-50000","50000以上"]
    all_txn["金额区间"] = pd.cut(all_txn["交易金额"], bins=bins, labels=labels, right=False, include_lowest=True)

    save_df_auto_width(all_txn, "所有人-合并交易流水", index=False, engine="openpyxl")
    print("✅ 已导出 所有人-合并交易流水.xlsx")
    return all_txn

# ------------------------------------------------------------------
# ⑦   单人资产 / 对手分析
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

    summary = pd.DataFrame([{
        "交易次数": len(df),
        "交易金额": df["交易金额"].sum(skipna=True),
        "流出额": out_df["交易金额"].sum(skipna=True),
        "流入额": in_df["交易金额"].sum(skipna=True),
        "单笔最大支出": out_df["交易金额"].max(skipna=True),
        "单笔最大收入": in_df["交易金额"].max(skipna=True),
        "净流入": in_df["交易金额"].sum(skipna=True) - out_df["交易金额"].sum(skipna=True),
        "最后交易时间": df["交易时间"].max(),
        "0-2千次数": counts.get("2000以下", 0),
        "2千-5千次数": counts.get("2000-5000", 0),
        "5千-2万次数": counts.get("5000-20000", 0),
        "2万-5万次数": counts.get("20000-50000", 0),
        "5万以上次数": counts.get("50000以上", 0),
    }])
    save_df_auto_width(summary, f"{prefix}0{person}-资产分析", index=False, engine="openpyxl")

    cash = df[(df["现金标志"].astype(str).str.contains("现", na=False) 
               | (pd.to_numeric(df["现金标志"], errors="coerce") == 1) 
               | df["交易类型"].astype(str).str.contains("柜面|现", na=False))
                & (pd.to_numeric(df["交易金额"], errors="coerce") >= 10_000)]
    save_df_auto_width(cash, f"{prefix}1{person}-存取现1万以上", index=False, engine="openpyxl")

    big = df[pd.to_numeric(df["交易金额"], errors="coerce") >= 500_000]
    save_df_auto_width(big, f"{prefix}1{person}-大额资金50万以上", index=False, engine="openpyxl")

    src = df.copy()
    src["is_in"] = src["借贷标志"] == "进"
    src["signed_amt"] = pd.to_numeric(src["交易金额"], errors="coerce") * src["is_in"].map({True: 1, False: -1})
    src["in_amt"] = pd.to_numeric(src["交易金额"], errors="coerce").where(src["is_in"], 0)
    src = (src.groupby("交易对方姓名", dropna=False)
           .agg(交易金额=("交易金额","sum"),
                交易次数=("交易金额","size"),
                流入额=("in_amt","sum"),
                净流入=("signed_amt","sum"),
                单笔最大收入=("in_amt","max"))
           .reset_index())
    total = src["流入额"].sum()
    src["流入比%"] = src["流入额"] / total * 100 if total else 0
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
    summ = (d.groupby(["查询对象","交易对方姓名"], dropna=False)
              .agg(交易次数=("交易金额","size"),
                   交易金额=("abs_amt","sum"),
                   万元以上交易次数=("gt10k","sum"),
                   净收入=("signed_amt","sum"),
                   转入笔数=("is_in","sum"),
                   转入金额=("in_amt","sum"),
                   转出笔数=("is_in", lambda x: (~x).sum()),
                   转出金额=("out_amt","sum"))
              .reset_index()
              .rename(columns={"查询对象":"姓名","交易对方姓名":"对方姓名"}))
    total = summ.groupby("姓名")["交易金额"].transform("sum")
    summ["交易占比%"] = np.where(total>0, summ["交易金额"] / total * 100, 0)
    summ.sort_values(["姓名","交易金额"], ascending=[True, False], inplace=True)
    save_df_auto_width(summ, f"{prefix}2{person}-交易对手分析", index=False, engine="openpyxl")
    comp = summ[summ["对方姓名"].astype(str).str.contains("公司", na=False)]
    save_df_auto_width(comp, f"{prefix}3{person}-与公司相关交易频次分析", index=False, engine="openpyxl")

# ------------------------------------------------------------------
# ⑧   GUI
# ------------------------------------------------------------------
def create_gui():
    root = tk.Tk()
    root.title("温岭纪委交易流水批量分析工具")
    root.minsize(780, 560)
    ttk.Label(root, text="温岭纪委交易流水批量分析工具", font=("仿宋", 20, "bold")).grid(row=0, column=0, columnspan=3, pady=(15, 0))
    ttk.Label(root, text="© 温岭纪委六室 单柳昊", font=("微软雅黑", 9)).grid(row=1, column=0, columnspan=3, pady=(0, 15))

    ttk.Label(root, text="工作目录:").grid(row=2, column=0, sticky="e", padx=8, pady=8)
    path_var = tk.StringVar(value=str(Path.home()))
    ttk.Entry(root, textvariable=path_var, width=60).grid(row=2, column=1, sticky="we", padx=5, pady=8)
    ttk.Button(
        root,
        text="浏览...",
        command=lambda: path_var.set(filedialog.askdirectory(title="选择工作目录") or path_var.get()),
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
