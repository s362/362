#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
交易流水批量分析工具 GUI   v6-plus (refactor + 线下银行扩展 + 通信联动修复)
Author  : 温岭纪委六室 单柳昊   （2025-08-05 修订）
重构者  : （效率优化版 2025-08-28）
扩展者  : （线下农行/建行接入 2025-09-09）
联动者  : （通信联动 2025-10-16，号码→姓名/职务回写通信，再以通信姓名→银行“对方职务”）
修复者  : （通讯录列识别&号码清洗修复 2025-10-16 增强版）

要点：
- 通讯录：手机号优先识别；表头自动探测；“职务”= 实际工作单位-（职务 / 职务或岗位 / 岗位）；号码非空即入库
- 通信：以【对方号码】（候选列）↔ 通讯录【手机号】匹配 → 回写 通信.姓名、通信.职务
- 银行：以 通信.姓名 ↔ 银行.交易对方姓名 匹配 → 写入 银行.对方职务
"""

import sys, os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading, warnings, builtins, datetime, re
from pathlib import Path
from functools import lru_cache
from typing import Optional, List, Dict, Any, Tuple

import pandas as pd
import numpy as np
from chinese_calendar import is_holiday, is_workday
from decimal import Decimal, InvalidOperation

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

# ===== 全局映射 =====
# 通讯录：手机号 -> (姓名, 职务)   # 职务 = 实际工作单位-（职务/职务或岗位/岗位）
CONTACT_PHONE_TO_NAME_TITLE: Dict[str, Tuple[str, str]] = {}
# 通信：姓名 -> 职务（仅来自号码匹配成功的通信记录）
CALLLOG_NAME_TO_TITLE: Dict[str, str] = {}

# ------------------------------------------------------------------
# 基础工具
# ------------------------------------------------------------------
SKIP_HEADER_KEYWORDS = ["反洗钱-电子账户交易明细","信用卡消费明细"]

@lru_cache(maxsize=None)
def _should_skip_special_cached(path_str: str) -> Optional[str]:
    p = Path(path_str)
    try:
        head = pd.read_excel(p, header=None, nrows=3)
        for kw in SKIP_HEADER_KEYWORDS:
            if head.astype(str).apply(lambda col: col.astype(str).str.contains(kw, na=False)).any().any():
                return kw
        return None
    except Exception:
        return None

def should_skip_special(p: Path) -> Optional[str]:
    return _should_skip_special_cached(str(p))

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
    """
    ★ 修改点（覆盖保存）：
    - 保存前若目标文件已存在，则先删除；再用 mode='w' 写入，确保覆盖。
    """
    if OUT_DIR is not None:
        filename = OUT_DIR / filename
    filename = Path(filename).with_suffix(".xlsx")
    filename.parent.mkdir(parents=True, exist_ok=True)

    # ★ 覆盖保存：先删除旧文件
    if filename.exists():
        try:
            filename.unlink()
        except Exception:
            pass

    df = df.replace(np.nan, "")
    if engine == "xlsxwriter":
        with pd.ExcelWriter(filename, engine="xlsxwriter", mode="w") as writer:  # ★ 明确写入模式
            df.to_excel(writer, sheet_name=sheet_name, index=index)
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                s = df[col].astype(str)
                width = max(min_width, min(max(s.map(len).max(), len(str(col))) + 2, max_width))
                ws.set_column(i, i, width)
    else:
        # openpyxl 流程也覆盖写
        with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:  # ★ 明确写入模式
            df.to_excel(writer, sheet_name=sheet_name, index=index)
        from openpyxl import load_workbook
        wb = load_workbook(filename)
        ws = wb[sheet_name]
        for col_cells in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells) + 2
            ws.column_dimensions[col_cells[0].column_letter].width = max(min_width, min(width, max_width)) + 5
        wb.save(filename)

def _iter_builtin_contacts_files() -> List[Path]:
    """
    返回内置通讯录文件列表（优先读同目录的 '内置-通讯录.xlsx'/'内置-通讯录.xls'）。
    同时兼容 PyInstaller 打包后的临时目录（sys._MEIPASS）。
    """
    candidates = ["内置-通讯录.xlsx", "内置-通讯录.xls"]
    base_dirs: List[Path] = []

    # 运行脚本所在目录
    try:
        base_dirs.append(Path(__file__).parent.resolve())
    except Exception:
        pass

    # PyInstaller 打包后的临时展开目录
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base_dirs.append(Path(sys._MEIPASS).resolve())

    # 当前工作目录（以防你把内置文件放在 cwd）
    try:
        base_dirs.append(Path.cwd().resolve())
    except Exception:
        pass

    out: List[Path] = []
    seen: set[str] = set()
    for b in base_dirs:
        for name in candidates:
            p = (b / name)
            try:
                rp = str(p.resolve())
            except Exception:
                rp = str(p)
            if p.exists() and rp not in seen:
                out.append(p)
                seen.add(rp)
    return out


# ------------------- 号码清洗（增强版） -------------------
_MOBILE_PAT = re.compile(r'(?:\+?86[-\s]?)?(1[3-9]\d{9})')
def normalize_phone_cell(x: Any) -> str:
    """
    强化版清洗规则：
    - 优先提取中国大陆手机号：1[3-9]开头11位（支持+86/86- 等前缀、各种分隔符、混写）
    - 若未命中，则保留数字并尝试：
        * 科学计数法/小数 .0
        * 仅数字时直接返回
    - 返回：首个手机号；若没有手机号但有其他数字，返回纯数字；实在没有返回空串
    """
    if x is None or (isinstance(x, float) and not np.isfinite(x)):
        return ""
    if isinstance(x, (int, np.integer)):
        s = str(int(x))
        m = _MOBILE_PAT.search(s)
        return m.group(1) if m else s
    if isinstance(x, (float, np.floating)):
        try:
            s = str(int(x))
            m = _MOBILE_PAT.search(s)
            return m.group(1) if m else s
        except Exception:
            pass

    s = str(x).strip().replace("\u00A0", " ")
    if s == "":
        return ""

    m = _MOBILE_PAT.search(s)
    if m:
        return m.group(1)

    if re.fullmatch(r"\d+(\.0+)?", s):
        return s.split(".")[0]

    if re.fullmatch(r"[0-9]+(\.[0-9]+)?[eE][+-]?[0-9]+", s):
        try:
            d = Decimal(s)
            q = d.quantize(Decimal(1))
            ss = format(q, "f")
            m2 = _MOBILE_PAT.search(ss)
            return m2.group(1) if m2 else ss
        except InvalidOperation:
            pass

    only_digits = re.sub(r"\D", "", s)
    if len(only_digits) >= 11:
        m3 = _MOBILE_PAT.search(only_digits)
        if m3:
            return m3.group(1)
    return only_digits

def str_to_weekday(date_val) -> str:
    dt = pd.to_datetime(date_val, errors="coerce")
    return "wrong" if pd.isna(dt) else ["星期一","星期二","星期三","星期四","星期五","星期六","星期日"][dt.weekday()]

def holiday_status(date_val) -> str:
    dt = pd.to_datetime(date_val, errors="coerce")
    if pd.isna(dt): return "wrong"
    d = dt.date()
    try:
        return "节假日" if is_holiday(d) else ("工作日" if is_workday(d) else "周末")
    except Exception:
        return "周末" if dt.weekday() >= 5 else "工作日"

# ---------- CSV ----------
def _read_csv_smart(p: Path, **kwargs) -> pd.DataFrame:
    enc_try = ["utf-8-sig", "gb18030", "utf-8", "cp936"]
    last_err: Optional[Exception] = None
    for enc in enc_try:
        try:
            return pd.read_csv(p, encoding=enc, **kwargs)
        except Exception as e:
            last_err = e
    raise last_err or RuntimeError(f"无法读取CSV: {p}")

def _person_from_people_csv(dirpath: Path) -> str:
    people = dirpath / "人员信息.csv"
    if not people.exists():
        return ""
    try:
        df = _read_csv_smart(people)
    except Exception:
        return ""
    for col in ["客户姓名", "姓名", "客户名称", "户名"]:
        if col in df.columns:
            ser = df[col].astype(str).str.strip()
            ser = ser[(ser != "") & (ser.str.lower() != "nan")]
            if not ser.empty:
                return ser.iloc[0][:10]
    name_pat = re.compile(r"客户(?:姓名|名称)\s*[:：]?\s*([^\s:：]{2,10})")
    vals = df.astype(str).replace("nan", "", regex=False).to_numpy().ravel().tolist()
    for val in vals:
        m = name_pat.search(val.strip())
        if m: return m.group(1)
    return ""

# ------------------------------------------------------------------
# 人名辅助
# ------------------------------------------------------------------
NAME_CANDIDATE_COLS: List[str] = ["账户名称","户名","账户名","账号名称","账号名","姓名","客户名称","查询对象"]

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
            header_idx = _header_row(fp)
            preview = pd.read_excel(fp, header=header_idx, nrows=5)
            if "账户名称" in preview.columns:
                s = preview["账户名称"].dropna()
                if not s.empty:
                    return str(s.iloc[0]).strip()
        except Exception:
            pass
    return ""

# ------------------------------------------------------------------
# 解析器（省略说明，保持你原逻辑不变）
# ------------------------------------------------------------------
@lru_cache(maxsize=None)
def _header_row(path: Path) -> int:
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

# ------------------------------------------------------------------
# CSV → 模板（保持原逻辑）
# ------------------------------------------------------------------
def csv_to_template(raw: pd.DataFrame, holder: str, feedback_unit: str) -> pd.DataFrame:
    if raw.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)
    try:
        df = raw.copy()
        df.columns = pd.Index(df.columns).astype(str).str.strip()
        n = len(df)
        def _S(default=""): return pd.Series([default]*n, index=df.index)
        def col(keys, default=""):
            if isinstance(keys, str): return df[keys] if keys in df else _S(default)
            for k in keys:
                if k in df: return df[k]
            return _S(default)
        def _to_str_no_sci(x):
            if pd.isna(x): return ""
            s = str(x).strip()
            if s.lower()=="nan": return ""
            if re.fullmatch(r"\d+\.0", s): return s[:-2]
            try:
                if isinstance(x, (int, np.integer)): return str(int(x))
                if isinstance(x, float): return f"{x:.0f}" if np.isfinite(x) else ""
                if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s): return f"{float(s):.0f}"
            except Exception: pass
            return s
        def _std_success(v):
            s = str(v).strip()
            if s in {"1","Y","y","是","成功","True","true"}: return "成功"
            if s in {"0","N","n","否","失败","False","false"}: return "失败"
            return "" if s.lower()=="nan" else s
        out = pd.DataFrame(index=df.index)
        out["本方账号"] = col(["交易账号","查询账户","本方账号","账号","账号/卡号","账号卡号"]).map(_to_str_no_sci)
        out["本方卡号"] = col(["交易卡号","查询卡号","本方卡号","卡号"]).map(_to_str_no_sci)
        out["查询账户"] = out["本方账号"]; out["查询卡号"]=out["本方卡号"]
        opp_no  = col(["交易对手账卡号","交易对手账号","对方账号","对方账户"]).map(_to_str_no_sci)
        opp_typ = col(["交易对方帐卡号类型","账号/卡号类型"], "")
        typ_s   = opp_typ.astype(str)
        is_card = typ_s.str.contains("卡", na=False) | typ_s.isin(["2","卡","卡号"])
        out["交易对方卡号"] = np.where(is_card, opp_no, ""); out["交易对方账户"]=np.where(is_card, "", opp_no)
        out["查询对象"] = holder or "未知"; out["反馈单位"]=feedback_unit or "未知"
        out["币种"] = col(["交易币种","币种","币别","货币"], "CNY").astype(str).replace(
            {"人民币":"CNY","人民币元":"CNY","RMB":"CNY","156":"CNY"}).fillna("CNY")
        out["交易金额"] = pd.to_numeric(col(["交易金额","金额","发生额"], 0), errors="coerce")
        out["账户余额"] = pd.to_numeric(col(["交易余额","余额","账户余额"], 0), errors="coerce")
        out["借贷标志"] = col(["收付标志",""], "")
        if "交易时间" in df.columns:
            tt = pd.to_datetime(df["交易时间"], errors="coerce")
            out["交易时间"] = np.where(tt.notna(), tt.dt.strftime("%Y-%m-%d %H:%M:%S"), df["交易时间"].astype(str))
        else:
            out["交易时间"] = _S("wrong")
        out["交易类型"] = col(["交易类型","业务种类","交易码"], "")
        out["交易流水号"] = col(["交易流水号","柜员流水号","流水号"], "")
        out["交易对方姓名"] = col(["对手户名","交易对手名称","对手方名称","对方户名","对方名称","对方姓名","收/付方名称"], " ")
        out["交易对方证件号码"] = col(["对手身份证号","对方证件号码"], " ")
        out["交易对手余额"] = pd.to_numeric(col(["对手交易余额"], ""), errors="coerce")
        out["交易对方账号开户行"] = col(["对手开户银行","交易对手行名","对方开户行","对方金融机构名称"], " ")
        out["交易摘要"] = col(["摘要说明","交易摘要","摘要","附言","用途"], " ")
        out["交易网点名称"] = col(["交易网点名称","交易机构","网点名称"], "")
        out["交易网点代码"] = col(["交易网点代码","机构号","网点代码"], "")
        out["日志号"] = col(["日志号"], ""); out["传票号"] = col(["传票号"], "")
        out["凭证种类"] = col(["凭证种类","凭证类型"], ""); out["凭证号"]=col(["凭证号","凭证序号"], "")
        out["现金标志"] = col(["现金标志"], ""); out["终端号"]=col(["终端号","渠道号"], "")
        succ = col(["交易是否成功","查询反馈结果"], ""); out["交易是否成功"]=succ.map(_std_success)
        out["交易发生地"] = col(["交易发生地","交易场所"], ""); out["商户名称"]=col(["商户名称"], ""); out["商户号"]=col(["商户号"], "")
        out["IP地址"]=col(["IP地址"], ""); out["MAC"]=col(["MAC地址","MAC"], ""); out["交易柜员号"]=col(["交易柜员号","柜员号","记账柜员"], "")
        try:
            beizhu = col(["备注","附言","说明"], "").astype(str); reason = col(["查询反馈结果原因"], "").astype(str)
            beizhu_clean = beizhu.where(~beizhu.str.lower().eq("nan"), ""); reason_clean = reason.where(~reason.str.lower().eq("nan"), "")
            out["备注"] = np.where(reason_clean!="", np.where(beizhu_clean!="" , beizhu_clean+"｜原因："+reason_clean, "原因："+reason_clean), beizhu_clean)
        except Exception:
            out["备注"] = _S("wrong")
        return out.reindex(columns=TEMPLATE_COLS, fill_value="")
    except Exception as e:
        print(f"❌ CSV转模板异常：{e}")
        return pd.DataFrame(columns=TEMPLATE_COLS)

# ===============================
# 各银行解析（保持原实现，略去注释）
# ===============================
def tl_to_template(raw) -> pd.DataFrame:
    if isinstance(raw, dict):
        frames=[]
        for sheet_name, df_sheet in raw.items():
            one = tl_to_template(df_sheet)
            if not one.empty:
                one.insert(0,"__sheet__",sheet_name); frames.append(one)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=TEMPLATE_COLS)
    df: pd.DataFrame = raw
    if df.empty: return pd.DataFrame(columns=TEMPLATE_COLS)
    def col_multi(keys, default=""):
        for k in keys:
            if k in df: return df[k]
        return pd.Series([default]*len(df), index=df.index)
    out = pd.DataFrame(index=df.index)
    out["本方账号"] = col_multi(["客户账号","账号","本方账号"], "wrong")
    out["查询账户"] = out["本方账号"]; out["反馈单位"]="泰隆银行"
    out["查询对象"] = col_multi(["账户名称","户名","客户名称"], "wrong")
    out["币种"] = col_multi(["币种","货币","币别"]).replace("156","CNY").replace("人民币元","CNY").replace("人民币","CNY").fillna("CNY")
    out["借贷标志"] = col_multi(["借贷标志","借贷方向","借贷"], "")
    debit  = pd.to_numeric(col_multi(["借方发生额","借方发生金额"], 0), errors="coerce")
    credit = pd.to_numeric(col_multi(["贷方发生额","贷方发生金额"], 0), errors="coerce")
    out["交易金额"] = debit.fillna(0).where(debit.ne(0), credit)
    out["账户余额"] = pd.to_numeric(col_multi(["账户余额","余额"], 0), errors="coerce")
    dates = col_multi(["交易日期","原交易日期","会计日期"]).astype(str)
    raw_times = col_multi(["交易时间","原交易时间","时间"]).astype(str).str.strip()
    def _tidy_time(s: str) -> str:
        if re.fullmatch(r"0+(\.0+)?", s): return ""
        if s.count(".") >= 2:
            p = s.split(".")
            if len(p[0])==2 and len(p[1])==2 and len(p[2])==2: return ".".join(p[:3])
        return s
    def _clean_time(s: str) -> str:
        s=s.strip()
        if re.fullmatch(r"0+(\.0+)?", s): return ""
        if re.fullmatch(r"\d{1,9}", s): return s.zfill(9)[:6]
        return s
    times = raw_times.apply(lambda x:_clean_time(_tidy_time(x)))
    out["交易时间"] = [ _parse_dt(d,t,False) for d,t in zip(dates,times)]
    out["交易流水号"] = col_multi(["原柜员流水号","流水号"])
    out["交易类型"] = col_multi(["交易码","交易类型","业务种类"])
    out["交易对方姓名"] = col_multi(["对方户名","交易对手名称"], " ")
    out["交易对方账户"] = col_multi(["对方客户账号","对方账号"], " ")
    out["交易对方账号开户行"] = col_multi(["对方金融机构名称","对方开户行"], " ")
    out["交易摘要"] = col_multi(["摘要描述","摘要"], " ")
    out["交易网点代码"] = col_multi(["机构号","网点代码"], " ")
    out["终端号"] = col_multi(["渠道号","终端号"], " ")
    out["交易柜员号"] = col_multi(["柜员号"], " ")
    out["备注"] = col_multi(["备注","附言"], " ")
    out["凭证种类"] = col_multi(["凭证类型"], ""); out["凭证号"]=col_multi(["凭证序号"], "")
    return out.reindex(columns=TEMPLATE_COLS, fill_value="")

def mt_to_template(raw: pd.DataFrame) -> pd.DataFrame:
    if raw.empty: return pd.DataFrame(columns=TEMPLATE_COLS)
    header_idx=None
    for i,row in raw.iterrows():
        cells=row.astype(str).str.strip().tolist()
        if "时间" in cells and "账号卡号" in cells:
            header_idx=i;break
    if header_idx is None:
        for i,row in raw.iterrows():
            if row.astype(str).str.contains("序号").any():
                header_idx=i;break
    if header_idx is None: return pd.DataFrame(columns=TEMPLATE_COLS)
    holder=""
    name_inline=re.compile(r"客户(?:姓名|名称)\s*[:：]?\s*([^\s:：]{2,10})")
    for i in range(header_idx):
        vals=raw.iloc[i].astype(str).tolist()
        for j,cell in enumerate(vals):
            cs=cell.strip(); m=name_inline.match(cs)
            if m: holder=m.group(1); break
            if re.fullmatch(r"客户(?:姓名|名称)\s*[:：]?", cs):
                nxt=str(vals[j+1]).strip() if j+1<len(vals) else ""
                if nxt and nxt.lower()!="nan": holder=nxt; break
        if holder: break
    holder=holder or "未知"
    df=raw.iloc[header_idx+1:].copy(); df.columns=raw.iloc[header_idx].astype(str).str.strip()
    df.dropna(how="all", inplace=True); df.reset_index(drop=True, inplace=True)
    summary_mask = df.apply(lambda row: row.astype(str).str.contains(r"支出笔数|收入笔数|支出累计金额|收入累计金额").any(), axis=1)
    df=df[~summary_mask].copy()
    def col(c, default=""): return df[c] if c in df else pd.Series(default, index=df.index)
    out=pd.DataFrame(index=df.index)
    acct=col("账号卡号").astype(str).str.replace(r"\.0$","",regex=True)
    out["本方账号"]=acct; out["查询账户"]=acct; out["查询对象"]=holder; out["反馈单位"]="民泰银行"
    out["币种"]=col("币种").astype(str).replace("人民币","CNY").replace("人民币元","CNY").fillna("CNY")
    debit=pd.to_numeric(col("支出"), errors="coerce").fillna(0)
    credit=pd.to_numeric(col("收入"), errors="coerce").fillna(0)
    out["交易金额"]=credit.where(credit.gt(0), -debit)
    out["账户余额"]=pd.to_numeric(col("余额"), errors="coerce")
    out["借贷标志"]=np.where(credit.gt(0),"进","出")
    def _fmt_time(v:str)->str:
        v=str(v).strip()
        try: return datetime.datetime.strptime(v,"%Y%m%d %H:%M:%S").strftime("%Y-%m-%d %H:%M:%S")
        except Exception: return v or "wrong"
    out["交易时间"]=col("时间").astype(str).apply(_fmt_time)
    out["交易摘要"]=col("摘要"," "); out["交易流水号"]=col("柜员流水号").astype(str).str.strip()
    out["交易柜员号"]=col("记账柜员 ").astype(str).str.strip(); out["交易网点代码"]=col("记账机构").astype(str).str.strip()
    out["交易对方姓名"]=col("交易对手名称"," ").astype(str).str.strip()
    out["交易对方账户"]=col("交易对手账号"," ").astype(str).str.strip()
    out["交易对方账号开户行"]=col("交易对手行名"," ").astype(str).str.strip()
    out["终端号"]=col("交易渠道"); out["备注"]=col("附言"," ")
    return out.reindex(columns=TEMPLATE_COLS, fill_value="")

def rc_to_template(raw: pd.DataFrame, holder: str, is_old: bool) -> pd.DataFrame:
    if raw.empty: return pd.DataFrame(columns=TEMPLATE_COLS)
    def col(c, default=""): return raw[c] if c in raw else pd.Series([default]*len(raw), index=raw.index)
    out=pd.DataFrame(index=raw.index)
    out["本方账号"]=col("账号","wrong"); out["查询账户"]=out["本方账号"]
    out["交易金额"]=col("发生额") if is_old else col("交易金额")
    out["账户余额"]=col("余额") if is_old else col("交易余额")
    out["反馈单位"]="老农商银行" if is_old else "新农商银行"
    dates=col("交易日期").astype(str); times=col("交易时间").astype(str)
    out["交易时间"]=[_parse_dt(d,t,is_old) for d,t in zip(dates,times)]
    out["借贷标志"]=col("借贷标志")
    out["币种"]="CNY" if is_old else col("币种").replace("人民币","CNY").replace("人民币元","CNY")
    out["查询对象"]=holder
    out["交易对方姓名"]=col("对方姓名"," "); out["交易对方账户"]=col("对方账号"," ")
    out["交易网点名称"]=col("代理行机构号") if is_old else col("交易机构")
    out["交易摘要"]=col("备注") if is_old else col("摘要","wrong")
    return out.reindex(columns=TEMPLATE_COLS, fill_value="")

# ---- 农行线下 APSH ----
def _is_abc_offline_file(p: Path) -> bool:
    try: xls = pd.ExcelFile(p); return "APSH" in xls.sheet_names
    except Exception: return False

def _merge_abc_datetime(date_val, time_val) -> str:
    ds = re.sub(r"\D","", "" if date_val is None else str(date_val).strip())
    date_ts = pd.to_datetime(ds[:8], format="%Y%m%d", errors="coerce") if len(ds)>=8 else pd.to_datetime(date_val, errors="coerce")
    if pd.isna(date_ts): return "wrong"
    def to_hhmmss(t)->str:
        if t is None or (isinstance(t,float) and np.isnan(t)) or t=="" or pd.isna(t): return "000000"
        if isinstance(t,(int,np.integer,float,np.floating)):
            try:
                tf=float(t)
                if 0<=tf<1:
                    secs=int(round(tf*86400)); secs=0 if secs>=86400 else secs
                    h=secs//3600; m=(secs%3600)//60; s=secs%60
                    return f"{h:02d}{m:02d}{s:02d}"
                digits=re.sub(r"\D","",str(int(round(tf)))); return digits.zfill(6)[:6]
            except Exception: pass
        s=str(t).strip()
        if ":" in s or "." in s:
            tt=pd.to_datetime("2000-01-01 "+s.replace(":",":").replace(".",":"), errors="coerce")
            if pd.notna(tt): return tt.strftime("%H%M%S")
        digits=re.sub(r"\D","",s); return (digits.zfill(6)[:6]) if digits!="" else "000000"
    hhmmss=to_hhmmss(time_val)
    return f"{date_ts.strftime('%Y-%m-%d')} {hhmmss[:2]}:{hhmmss[2:4]}:{hhmmss[4:6]}"

def abc_offline_from_file(p: Path) -> pd.DataFrame:
    try:
        xls=pd.ExcelFile(p)
        if "APSH" not in xls.sheet_names: return pd.DataFrame(columns=TEMPLATE_COLS)
        df=xls.parse("APSH", header=0)
    except Exception: return pd.DataFrame(columns=TEMPLATE_COLS)
    if df.empty: return pd.DataFrame(columns=TEMPLATE_COLS)
    df.columns=pd.Index(df.columns).astype(str).str.strip()
    n=len(df); out=pd.DataFrame(index=df.index)
    out["本方账号"]=df.get("账号","")
    out["本方卡号"]=df.get("卡号","").astype(str).str.replace(r"\.0$","",regex=True)
    out["查询账户"]=out["本方账号"]; out["查询卡号"]=out["本方卡号"]
    holder=df.get("户名",""); holder = pd.Series([holder]*n,index=df.index) if not isinstance(holder,pd.Series) else holder
    out["查询对象"]=holder.fillna("").astype(str).str.strip().replace({"nan":""}).replace("","未知")
    out["反馈单位"]="农业银行"; out["币种"]="CNY"
    amt=pd.to_numeric(df.get("交易金额",0), errors="coerce"); out["交易金额"]=amt
    out["账户余额"]=pd.to_numeric(df.get("交易后余额",""), errors="coerce")
    out["借贷标志"]=np.where(amt>0,"进",np.where(amt<0,"出",""))
    dates=df.get("交易日期",""); times=df.get("交易时间","")
    out["交易时间"]=[_merge_abc_datetime(d,t) for d,t in zip(dates,times)]
    out["交易摘要"]=df.get("摘要","").astype(str); out["交易流水号"]=""
    out["交易类型"]=""
    out["交易对方姓名"]=df.get("对方户名"," ").astype(str)
    out["交易对方账户"]=df.get("对方账号"," ").astype(str)
    out["交易对方卡号"]=""
    out["交易对方证件号码"]=" "; out["交易对手余额"]=""
    out["交易对方账号开户行"]=df.get("对方开户行"," ").astype(str)
    out["交易网点名称"]=df.get("交易网点","").astype(str)
    out["交易网点代码"]=df.get("交易行号","").astype(str)
    out["日志号"]=""
    out["传票号"]=df.get("传票号","").astype(str)
    out["凭证种类"]=""
    out["凭证号"]=""
    out["现金标志"]=""
    out["终端号"]=df.get("交易渠道","").astype(str)
    out["交易是否成功"]=""
    out["交易发生地"]=""
    out["商户名称"]=""
    out["商户号"]=""
    out["IP地址"]=""
    out["MAC"]=""
    out["交易柜员号"]=""
    out["备注"]=""
    return out.reindex(columns=TEMPLATE_COLS, fill_value="")

# ---- 建行线下 交易明细 ----
def _is_ccb_offline_file(p: Path) -> bool:
    try:
        xls=pd.ExcelFile(p)
        if "交易明细" not in xls.sheet_names: return False
        df_head=xls.parse("交易明细", nrows=1)
        cols=set(map(str, df_head.columns))
        return {"客户名称","账号","交易日期","交易时间","交易金额"}.issubset(cols)
    except Exception: return False

def ccb_offline_from_file(p: Path) -> pd.DataFrame:
    try:
        xls=pd.ExcelFile(p)
        if "交易明细" not in xls.sheet_names: return pd.DataFrame(columns=TEMPLATE_COLS)
        df=xls.parse("交易明细", header=0)
    except Exception: return pd.DataFrame(columns=TEMPLATE_COLS)
    if df.empty: return pd.DataFrame(columns=TEMPLATE_COLS)
    df.columns=pd.Index(df.columns).astype(str).str.strip()
    out=pd.DataFrame(index=df.index)
    out["本方账号"]=df.get("账号","")
    out["本方卡号"]=df.get("交易卡号","").astype(str).str.replace(r"\.0$","",regex=True)
    out["查询账户"]=out["本方账号"]; out["查询卡号"]=out["本方卡号"]
    out["查询对象"]=df.get("客户名称","").astype(str).replace({"nan":""}).replace("","未知")
    out["反馈单位"]="建设银行"
    out["币种"]=df.get("币种","CNY").astype(str).replace({"人民币":"CNY","人民币元":"CNY","RMB":"CNY","156":"CNY"}).fillna("CNY")
    amt=pd.to_numeric(df.get("交易金额",0), errors="coerce"); out["交易金额"]=amt
    out["账户余额"]=pd.to_numeric(df.get("账户余额",""), errors="coerce")
    jd=df.get("借贷方向","").astype(str).str.strip()
    out["借贷标志"]=np.where(jd.str.contains("^贷",na=False)|jd.str.upper().isin(["贷","C","CR","CREDIT"]),"进",
                        np.where(jd.str.contains("^借",na=False)|jd.str.upper().isin(["借","D","DR","DEBIT"]),"出",
                                 np.where(amt>0,"进",np.where(amt<0,"出",""))))
    dates=df.get("交易日期",""); times=df.get("交易时间",""); times_str=pd.Series(times).astype(str).str.replace(r"\.0$","",regex=True)
    out["交易时间"]=[_parse_dt(d,t,False) for d,t in zip(dates,times_str)]
    out["交易摘要"]=df.get("摘要"," ").astype(str); out["交易类型"]=""
    out["交易流水号"]=df.get("交易流水号","").astype(str)
    out["交易对方姓名"]=df.get("对方户名"," ").astype(str)
    out["交易对方账户"]=df.get("对方账号"," ").astype(str)
    out["交易对方卡号"]=""
    out["交易对方证件号码"]=" "; out["交易对手余额"]=""
    out["交易对方账号开户行"]=df.get("对方行名"," ").astype(str)
    out["交易网点名称"]=df.get("交易机构名称","").astype(str)
    out["交易网点代码"]=df.get("交易机构号","").astype(str)
    out["交易柜员号"]=df.get("柜员号","").astype(str)
    out["终端号"]=df.get("交易渠道","").astype(str)
    ext=df.get("扩充备注","").astype(str).replace({"nan":""}); out["备注"]=ext
    out["现金标志"]=""; out["日志号"]=""; out["传票号"]=""
    out["凭证种类"]=""; out["凭证号"]=""
    out["交易是否成功"]=""
    out["交易发生地"]=""
    out["商户名称"]=df.get("商户名称","").astype(str)
    out["商户号"]=df.get("商户号","").astype(str)
    out["IP地址"]=df.get("IP地址","").astype(str)
    out["MAC"]=df.get("MAC地址","").astype(str)
    return out.reindex(columns=TEMPLATE_COLS, fill_value="")

# ------------------------------------------------------------------
# 通讯录读取（修复版）—— 自动探测表头；手机号提取+兜底扫描
# ------------------------------------------------------------------
CONTACT_NAME_COLS = ["姓名","联系人","人员姓名","姓名/名称"]
CONTACT_DEPT_KEYS = ["实际工作单位"]                     # 固定
CONTACT_TITLE_KEYS = ["职务","职务或岗位","岗位"]         # 新增“职务或岗位”

# 【关键修复】—— 扩充可作为“号码列”的关键字
CONTACT_PHONE_KEYS = [
    "号码","手机号","手机号码","联系电话","电话","联系方式","联系号码","移动电话","联系手机","联系电话（手机）","手机"
]

def _guess_header_row(xls: pd.ExcelFile, sheet_name: str, scan_rows: int = 30) -> int:
    df0 = xls.parse(sheet_name, header=None, nrows=scan_rows)
    for i, row in df0.iterrows():
        if row.astype(str).str.contains("姓名|号码|联系电话|电话|手机号|职务|岗位|实际工作单位|联系方式").any():
            return i
    return 0

def _compose_title(dept: Any, title: Any) -> str:
    d = "" if pd.isna(dept) else str(dept).strip()
    t = "" if pd.isna(title) else str(title).strip()
    if d and t: return f"{d}-{t}"
    if t: return t
    if d: return d
    return ""

def _find_first_col(df: pd.DataFrame, keys: List[str]) -> Optional[str]:
    for c in df.columns:
        cs = str(c).strip()
        for k in keys:
            if k in cs:
                return c
    return None

def _extract_mobile_from_row(row: pd.Series) -> str:
    """兜底：在整行文本里搜第一个大陆手机号"""
    text = " ".join(map(lambda v: "" if pd.isna(v) else str(v), row.values))
    m = _MOBILE_PAT.search(text)
    return m.group(1) if m else ""

def load_contacts_phone_map(root: Path) -> Dict[str, Tuple[str,str]]:
    # 固定提示
    print("正在读取通讯录......")

    def _is_in_out_dir(p: Path) -> bool:
        try:
            return OUT_DIR is not None and p.resolve().is_relative_to(OUT_DIR.resolve())
        except AttributeError:
            return OUT_DIR is not None and str(p.resolve()).startswith(str(OUT_DIR.resolve()))

    # ★ 1) 先加入“内置-通讯录.xlsx/.xls”（始终读取）
    builtin_files = _iter_builtin_contacts_files()
    if builtin_files:
        for bp in builtin_files:
            print(f"  • 使用内置通讯录：{bp.name}")

    # ★ 2) 再收集工作目录下普通通讯录（跳过已标注/输出目录）
    repo_files = [
        p for p in root.rglob("*通讯录*.xls*")
        if ("已标注" not in p.stem) and (not _is_in_out_dir(p))
    ]

    # 合并并去重（按真实路径）
    all_files: List[Path] = []
    seen: set[str] = set()
    for p in [*builtin_files, *repo_files]:
        try:
            rp = str(p.resolve())
        except Exception:
            rp = str(p)
        if rp not in seen:
            all_files.append(p)
            seen.add(rp)

    if not all_files:
        print("ℹ️ 未发现可用的通讯录。")
        return {}

    merged: Dict[str, Tuple[str,str]] = {}
    total_rows = 0
    total_with_phone = 0

    for p in all_files:
        try:
            xls = pd.ExcelFile(p)
        except Exception as e:
            print("❌ 通讯录载入失败", p.name, e); 
            continue

        for sht in xls.sheet_names:
            try:
                # 1) 探测表头
                hdr_row = _guess_header_row(xls, sht, 30)

                # 2) 先读一遍拿到列名
                df0 = xls.parse(sht, header=hdr_row)
                df0.columns = pd.Index(df0.columns).astype(str).str.strip()

                # 3) 找关键列名
                def _find_col(keys: List[str]) -> Optional[str]:
                    for c in df0.columns:
                        cs = str(c).strip()
                        for k in keys:
                            if k in cs:
                                return c
                    return None

                name_col  = _find_col(CONTACT_NAME_COLS)
                phone_col = _find_col(CONTACT_PHONE_KEYS)
                dept_col  = _find_col(CONTACT_DEPT_KEYS)
                title_col = _find_col(CONTACT_TITLE_KEYS)

                # 4) 强制把【号码列】按字符串重读；若找不到号码列，也先按普通读，后续兜底扫描
                dtype_kw = {phone_col: str} if phone_col else None
                df = xls.parse(sht, header=hdr_row, dtype=dtype_kw)
                df.columns = pd.Index(df.columns).astype(str).str.strip()

                total_rows += len(df.index)

                nm_ser   = df[name_col].astype(str).str.strip() if name_col else pd.Series([""]*len(df))
                dept_ser = df[dept_col] if dept_col else pd.Series([""]*len(df))
                titl_ser = df[title_col] if title_col else pd.Series([""]*len(df))

                if phone_col:
                    raw_phone = df[phone_col]
                    phone_ser = raw_phone.map(normalize_phone_cell)
                else:
                    # 没有显式号码列：对整行做手机号扫描
                    phone_ser = df.apply(_extract_mobile_from_row, axis=1)

                # sheet 级统计
                sheet_phones = phone_ser.astype(bool).sum()
                total_with_phone += int(sheet_phones)
                print(f"  • 通讯录 {p.name} / {sht}: 行数 {len(df)}, 命中手机号 {int(sheet_phones)}")

                for nm, ph, dp, tt in zip(nm_ser, phone_ser, dept_ser, titl_ser):
                    if not ph:
                        continue
                    job = _compose_title(dp, tt)  # 实际工作单位-（职务/职务或岗位/岗位）
                    if ph not in merged:
                        merged[ph] = (nm, job)
                    else:
                        old_nm, old_job = merged[ph]
                        merged[ph] = (old_nm or nm, old_job or job)

            except Exception as e:
                print("❌ 通讯录解析失败", f"{p.name}->{sht}", e)

    print(f"✅ 通讯录号码映射加载完成：{len(merged)} 条（扫描行数 {total_rows}；含手机号 {total_with_phone}）。")
    # 不输出样例
    return merged


# ------------------------------------------------------------------
# 通信：号码匹配通讯录 → 回写通信“姓名”“职务”；输出 通信“姓名→职务”
# ------------------------------------------------------------------
CALLLOG_PHONE_COL_CANDS = ["对方号码"]
CALLLOG_NAME_COL_CANDS  = ["姓名","对方姓名","联系人","名称","客户名称","户名"]

def _find_col_any(df: pd.DataFrame, cands: List[str]) -> Optional[str]:
    for c in map(str, df.columns):
        sc = c.strip()
        for key in cands:
            if key in sc:
                return c
    return None

def _enrich_one_comm_df(df: pd.DataFrame, phone_to_name_title: Dict[str, Tuple[str,str]]) -> Tuple[pd.DataFrame, Dict[str,str]]:
    if df is None or df.empty: return pd.DataFrame(), {}
    d = df.copy(); d.columns = pd.Index(d.columns).astype(str).str.strip()
    phone_col = _find_col_any(d, CALLLOG_PHONE_COL_CANDS)
    if not phone_col:
        return d, {}
    phones = d[phone_col].map(normalize_phone_cell)

    names = []
    titles = []
    for ph in phones:
        nm, job = phone_to_name_title.get(ph, ("",""))
        names.append(nm)
        titles.append(job)
    names  = pd.Series(names, index=d.index)
    titles = pd.Series(titles, index=d.index)

    name_col_existing = _find_col_any(d, CALLLOG_NAME_COL_CANDS)
    if name_col_existing:
        d["姓名"] = np.where(names!="", names, d[name_col_existing].astype(str).str.strip())
    else:
        d["姓名"] = names

    d["职务"] = titles  # 号码命中则有值，否则为空

    tmp = d[["姓名","职务"]].copy()
    tmp = tmp[(tmp["姓名"]!="") & (~tmp["姓名"].str.lower().eq("nan")) & (tmp["职务"]!="")]
    map_name_title: Dict[str,str] = {}
    for name, sub in tmp.groupby("姓名"):
        uniq = list(dict.fromkeys(sub["职务"].astype(str).tolist()))
        map_name_title[name] = "、".join(x for x in uniq if x)
    return d, map_name_title

def load_and_enrich_communications(root: Path, phone_to_name_title: Dict[str, Tuple[str,str]]) -> Dict[str,str]:
    if not phone_to_name_title:
        print("ℹ️ 未能从通讯录生成号码映射，跳过通信标注。")
        return {}

    def _is_in_out_dir(p: Path) -> bool:
        try:
            return OUT_DIR is not None and p.resolve().is_relative_to(OUT_DIR.resolve())
        except AttributeError:
            # Python<3.9 没有 is_relative_to
            return OUT_DIR is not None and str(p.resolve()).startswith(str(OUT_DIR.resolve()))

    # ★ 仅读取原始“通信”文件；跳过“已标注”与输出目录里的文件
    files = [
        p for p in root.rglob("*.xlsx")
        if ("通信" in p.stem or "通信" in p.name)
        and ("已标注" not in p.stem)
        and (not _is_in_out_dir(p))
    ]

    if not files:
        print("ℹ️ 未发现文件名包含“通信”的 .xlsx。")
        return {}
    out_all: Dict[str,str] = {}
    for p in files:
        print(f"📞 通信匹配：{p.name} ...")
        try:
            xls = pd.ExcelFile(p)
        except Exception as e:
            print("❌ 通信文件载入失败", p.name, e); continue
        frames = []; name_map_file: Dict[str,str] = {}
        for sht in xls.sheet_names:
            try:
                df0 = xls.parse(sheet_name=sht, header=0)
            except Exception as e:
                print("❌ 通信解析失败", f"{p.name}->{sht}", e); continue
            enriched, name_map = _enrich_one_comm_df(df0, phone_to_name_title)
            if not enriched.empty:
                enriched.insert(0,"__来源sheet__",sht); frames.append(enriched)
            for k,v in name_map.items():
                if k in name_map_file and name_map_file[k]:
                    exist = name_map_file[k].split("、")
                    add = [x for x in v.split("、") if x not in exist]
                    name_map_file[k] = "、".join(exist + add)
                else:
                    name_map_file[k] = v
        if frames:
            merged = pd.concat(frames, ignore_index=True)
            save_df_auto_width(merged, Path("通信-已标注")/f"{p.stem}-已标注", index=False, engine="openpyxl")
            print(f"✅ 通信已标注导出：{p.stem}-已标注.xlsx")
        for k,v in name_map_file.items():
            if k in out_all and out_all[k]:
                exist = out_all[k].split("、")
                add = [x for x in v.split("、") if x not in exist]
                out_all[k] = "•".join(exist)  # 不重要；合并策略同上
                out_all[k] = "、".join(exist + add)
            else:
                out_all[k] = v
    print(f"✅ 通信姓名映射生成 {len(out_all)} 条。")
    return out_all

# ------------------------------------------------------------------
# 合并全部流水（在导出前把“对方职务”接上）
# ------------------------------------------------------------------
def merge_all_txn(root_dir: str) -> pd.DataFrame:
    root = Path(root_dir).expanduser().resolve()

    # 1) 通讯录：手机号 -> (姓名, 职务)
    global CONTACT_PHONE_TO_NAME_TITLE, CALLLOG_NAME_TO_TITLE
    CONTACT_PHONE_TO_NAME_TITLE = load_contacts_phone_map(root)

    # 2) 通信：以号码匹配通讯录，回写通信中的“姓名”“职务”，并汇总“姓名→职务”
    CALLLOG_NAME_TO_TITLE = load_and_enrich_communications(root, CONTACT_PHONE_TO_NAME_TITLE)

    # 3) 收集银行与其他来源文件
    china_files = [p for p in root.rglob("*-*-交易流水.xls*")]
    all_excel = list(root.rglob("*.xls*"))
    rc_candidates = [p for p in all_excel if "农商行" in p.as_posix()]
    pattern_old = re.compile(r"老\s*[账帐]\s*(?:号|户)")
    old_rc = [p for p in rc_candidates if pattern_old.search(p.stem)]
    new_rc = [p for p in rc_candidates if p not in old_rc]
    tl_files = [p for p in all_excel if "泰隆" in p.as_posix()]
    mt_files = [p for p in all_excel if "民泰" in p.as_posix()]
    abc_offline_files = [p for p in all_excel if _is_abc_offline_file(p)]
    ccb_offline_files = [p for p in all_excel if _is_ccb_offline_file(p)]
    csv_txn_files = [p for p in root.rglob("交易明细信息.csv")]

    print(f"✅ 网上银行 {len(china_files)}，老农商 {len(old_rc)}，新农商 {len(new_rc)}，泰隆 {len(tl_files)}，民泰 {len(mt_files)}，农行线下 {len(abc_offline_files)}，建行线下 {len(ccb_offline_files)}，CSV {len(csv_txn_files)}；通信映射 {len(CALLLOG_NAME_TO_TITLE)} 条。")

    dfs: List[pd.DataFrame] = []

    # 网银标准
    for p in china_files:
        print(f"正在处理 {p.name} ...")
        try:
            df = pd.read_excel(p, dtype={"查询卡号":str,"查询账户":str,"交易对方证件号码":str,"本方账号":str,"本方卡号":str})
            df["来源文件"] = p.name
            dfs.append(df)
        except Exception as e:
            print("❌", p.name, e)

    # 农商行
    for p in old_rc + new_rc:
        print(f"正在处理 {p.name} ...")
        kw = should_skip_special(p)
        if kw:
            print(f"⏩ 跳过【{p.name}】：表头含“{kw}”"); continue
        raw = _read_raw(p)
        holder = extract_holder_from_df(raw) or holder_from_folder(p.parent) or fallback_holder_from_path(p)
        is_old = p in old_rc
        df = rc_to_template(raw, holder, is_old)
        if not df.empty:
            df["来源文件"] = p.name; dfs.append(df)

    # 泰隆
    for p in tl_files:
        if "开户" in p.stem: continue
        print(f"正在处理 {p.name} ...")
        try:
            xls = pd.ExcelFile(p)
        except Exception as e:
            print("❌", f"{p.name} 载入失败", e); continue
        try:
            header_idx = _header_row(p)
        except Exception:
            header_idx = 0
        xls_dict={}
        for sht in xls.sheet_names:
            try:
                df_sheet = xls.parse(sheet_name=sht, header=header_idx)
                xls_dict[sht]=df_sheet
            except Exception as e:
                print("❌", f"{p.name}->{sht}", e)
        df = tl_to_template(xls_dict)
        if not df.empty:
            df["来源文件"]=p.name; dfs.append(df)

    # 民泰
    for p in mt_files:
        print(f"正在处理 {p.name} ...")
        raw = _read_raw(p); df = mt_to_template(raw)
        if not df.empty:
            df["来源文件"]=p.name; dfs.append(df)

    # 农行线下
    for p in abc_offline_files:
        print(f"正在处理 {p.name} ...")
        try:
            df=abc_offline_from_file(p)
            if not df.empty:
                df["来源文件"]=p.name; dfs.append(df)
        except Exception as e:
            print("❌ 农行线下解析失败", p.name, e)

    # 建行线下
    for p in ccb_offline_files:
        print(f"正在处理 {p.name} ...")
        try:
            df=ccb_offline_from_file(p)
            if not df.empty:
                df["来源文件"]=p.name; dfs.append(df)
        except Exception as e:
            print("❌ 建行线下解析失败", p.name, e)

    # CSV
    for p in csv_txn_files:
        print(f"正在处理 {p.name} ...")
        try:
            raw_csv = _read_csv_smart(p)
        except Exception as e:
            print("❌ 无法读取CSV", p.name, e); continue
        holder = _person_from_people_csv(p.parent) or holder_from_folder(p.parent) or fallback_holder_from_path(p)
        feedback_unit = p.parent.name
        try:
            df = csv_to_template(raw_csv, holder, feedback_unit)
        except Exception as e:
            print("❌ CSV转模板失败", p.name, e); continue
        if not df.empty:
            df["来源文件"]=p.name; dfs.append(df)

    print("文件读取完成，正在整合……")
    if not dfs:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    all_txn = pd.concat(dfs, ignore_index=True)

    # 去重：流水号+时间+金额
    all_txn["交易金额"] = pd.to_numeric(all_txn["交易金额"], errors="coerce").round(2)
    before=len(all_txn)
    all_txn = all_txn.drop_duplicates(subset=["交易流水号","交易时间","交易金额"], keep="first").reset_index(drop=True)
    removed=before-len(all_txn)
    if removed: print(f"🧹 去重 {removed} 条。")

    # 排序/序号
    ts = pd.to_datetime(all_txn["交易时间"], errors="coerce")
    all_txn.insert(0,"__ts__",ts)
    all_txn.sort_values("__ts__", inplace=True, kind="mergesort")
    all_txn["序号"] = range(1, len(all_txn)+1)
    all_txn.drop(columns="__ts__", inplace=True)

    # 标准化、分箱、星期/节假日
    all_txn["借贷标志"]=all_txn["借贷标志"].apply(lambda x: "出" if str(x).strip() in {"1","借","D"} else ("进" if str(x).strip() in {"2","贷","C"} else str(x)))
    bins=[-np.inf,2000,5000,20000,50000,np.inf]; labels=["2000以下","2000-5000","5000-20000","20000-50000","50000以上"]
    all_txn["金额区间"]=pd.cut(pd.to_numeric(all_txn["交易金额"], errors="coerce"), bins=bins, labels=labels, right=False, include_lowest=True)
    weekday_map={0:"星期一",1:"星期二",2:"星期三",3:"星期四",4:"星期五",5:"星期六",6:"星期日"}
    wk = pd.Series(index=all_txn.index, dtype=object); mask=ts.notna()
    wk.loc[mask]=ts.dt.weekday.map(weekday_map); wk.loc[~mask]="wrong"; all_txn["星期"]=wk
    dates=ts.dt.date; status=pd.Series(index=all_txn.index, dtype=object)
    unique_dates=pd.unique(dates[mask])
    @lru_cache(maxsize=None)
    def _day_status(d)->str:
        try: return "节假日" if is_holiday(d) else ("工作日" if is_workday(d) else "周末")
        except Exception:
            dd=datetime.datetime.combine(d, datetime.time())
            return "周末" if dd.weekday()>=5 else "工作日"
    if len(unique_dates):
        mapd={d:_day_status(d) for d in unique_dates}; status.loc[mask]=dates.loc[mask].map(mapd)
    status.loc[~mask]="wrong"; all_txn["节假日"]=status

    # —— 对方职务：优先“通信姓名→职务”，★ 若无则回退“通讯录姓名→职务”
    #   a) 通信姓名→职务（已有）
    name_from_comm = CALLLOG_NAME_TO_TITLE or {}

    #   b) 通讯录姓名→职务（将手机号映射反向：姓名 -> 职务），取第一个非空职务
    contacts_name_to_title: Dict[str, str] = {}
    for ph, (nm, job) in CONTACT_PHONE_TO_NAME_TITLE.items():
        if not nm:
            continue
        if nm not in contacts_name_to_title:
            contacts_name_to_title[nm] = job or ""
        else:
            # 空则用新的非空覆盖
            if not contacts_name_to_title[nm] and job:
                contacts_name_to_title[nm] = job

    #   c) 最终映射：先用通信，缺失再用通讯录
    final_title_by_name = name_from_comm.copy()
    for nm, job in contacts_name_to_title.items():
        if nm not in final_title_by_name or not final_title_by_name[nm]:
            final_title_by_name[nm] = job

    all_txn["对方职务"] = (
        all_txn["交易对方姓名"]
        .map(final_title_by_name)
        .fillna("")
    )

    save_df_auto_width(all_txn, "所有人-合并交易流水", index=False, engine="openpyxl")
    print("✅ 已导出：所有人-合并交易流水.xlsx")
    return all_txn

# ------------------------------------------------------------------
# 分析（对方职务直接使用合并结果）
# ------------------------------------------------------------------
def analysis_txn(df: pd.DataFrame) -> None:
    if df.empty: return
    df=df.copy()
    df["交易时间"]=pd.to_datetime(df["交易时间"], errors="coerce")
    df["交易金额"]=pd.to_numeric(df["交易金额"], errors="coerce")
    person=df["查询对象"].iat[0] or "未知"
    prefix=f"{person}/"

    out_df=df[df["借贷标志"]=="出"]; in_df=df[df["借贷标志"]=="进"]; counts=df["金额区间"].value_counts()
    summary=pd.DataFrame([{
        "交易次数":len(df),
        "交易金额":df["交易金额"].sum(skipna=True),
        "流出额":out_df["交易金额"].sum(skipna=True),
        "流入额":in_df["交易金额"].sum(skipna=True),
        "单笔最大支出":out_df["交易金额"].max(skipna=True),
        "单笔最大收入":in_df["交易金额"].max(skipna=True),
        "净流入":in_df["交易金额"].sum(skipna=True)-out_df["交易金额"].sum(skipna=True),
        "最后交易时间":df["交易时间"].max(),
        "0-2千次数":counts.get("2000以下",0),
        "2千-5千次数":counts.get("2000-5000",0),
        "5千-2万次数":counts.get("5000-20000",0),
        "2万-5万次数":counts.get("20000-50000",0),
        "5万以上次数":counts.get("50000以上",0),
    }])
    save_df_auto_width(summary, f"{prefix}0{person}-资产分析", index=False, engine="openpyxl")

    cash = df[(df["现金标志"].astype(str).str.contains("现", na=False)
               | (pd.to_numeric(df["现金标志"], errors="coerce")==1)
               | df["交易类型"].astype(str).str.contains("柜面|现", na=False))
              & (pd.to_numeric(df["交易金额"], errors="coerce")>=10_000)]
    save_df_auto_width(cash, f"{prefix}1{person}-存取现1万以上", index=False, engine="openpyxl")

    big = df[pd.to_numeric(df["交易金额"], errors="coerce")>=500_000]
    save_df_auto_width(big, f"{prefix}1{person}-大额资金50万以上", index=False, engine="openpyxl")

    src=df.copy()
    src["is_in"]=src["借贷标志"]=="进"
    src["signed_amt"]=pd.to_numeric(src["交易金额"], errors="coerce")*src["is_in"].map({True:1,False:-1})
    src["in_amt"]=pd.to_numeric(src["交易金额"], errors="coerce").where(src["is_in"],0)
    src=(src.groupby("交易对方姓名", dropna=False)
         .agg(交易金额=("交易金额","sum"),
              交易次数=("交易金额","size"),
              流入额=("in_amt","sum"),
              净流入=("signed_amt","sum"),
              单笔最大收入=("in_amt","max"))
         .reset_index())
    total=src["流入额"].sum()
    src["流入比%"]=src["流入额"]/total*100 if total else 0
    # 并入对方职务（已包含通信优先+通讯录回退）
    name_to_title = (df[["交易对方姓名","对方职务"]].dropna().drop_duplicates().set_index("交易对方姓名")["对方职务"].to_dict())
    src.insert(1,"对方职务", src["交易对方姓名"].map(name_to_title).fillna(""))
    src.sort_values("流入额", ascending=False, inplace=True)
    save_df_auto_width(src, f"{prefix}1{person}-资金来源分析", index=False, engine="openpyxl")

def make_partner_summary(df: pd.DataFrame) -> None:
    if df.empty: return
    person=df["查询对象"].iat[0] or "未知"; prefix=f"{person}/"
    d=df.copy()
    d["交易金额"]=pd.to_numeric(d["交易金额"], errors="coerce").fillna(0)
    d["is_in"]=d["借贷标志"]=="进"
    d["abs_amt"]=d["交易金额"].abs()
    d["signed_amt"]=d["交易金额"]*d["is_in"].map({True:1,False:-1})
    d["in_amt"]=d["交易金额"].where(d["is_in"],0)
    d["out_amt"]=d["交易金额"].where(~d["is_in"],0)
    d["gt10k"]=(d["abs_amt"]>=10_000).astype(int)
    summ=(d.groupby(["查询对象","交易对方姓名"], dropna=False)
            .agg(交易次数=("交易金额","size"),
                 交易金额=("abs_amt","sum"),
                 万元以上交易次数=("gt10k","sum"),
                 净收入=("signed_amt","sum"),
                 转入笔数=("is_in","sum"),
                 转入金额=("in_amt","sum"),
                 转出笔数=("is_in", lambda x:(~x).sum()),
                 转出金额=("out_amt","sum"))
            .reset_index()
            .rename(columns={"查询对象":"姓名","交易对方姓名":"对方姓名"}))
    name_to_title=(d[["交易对方姓名","对方职务"]].drop_duplicates().set_index("交易对方姓名")["对方职务"].to_dict())
    summ.insert(2,"对方职务", summ["对方姓名"].map(name_to_title).fillna(""))
    total=summ.groupby("姓名")["交易金额"].transform("sum")
    summ["交易占比%"]=np.where(total>0, summ["交易金额"]/total*100, 0)
    summ.sort_values(["姓名","交易金额"], ascending=[True,False], inplace=True)
    save_df_auto_width(summ, f"{prefix}2{person}-交易对手分析", index=False, engine="openpyxl")

    comp=summ[summ["对方姓名"].astype(str).str.contains("公司", na=False)]
    save_df_auto_width(comp, f"{prefix}3{person}-与公司相关交易频次分析", index=False, engine="openpyxl")

# ------------------------------------------------------------------
# GUI
# ------------------------------------------------------------------
def create_gui():
    root = tk.Tk()
    root.title("温岭纪委交易流水批量分析工具")
    root.minsize(820, 600)

    ttk.Label(root, text="温岭纪委交易流水批量分析工具", font=("仿宋", 20, "bold")).grid(row=0, column=0, columnspan=3, pady=(15, 0))
    ttk.Label(root, text="© 温岭纪委六室 单柳昊", font=("微软雅黑", 9)).grid(row=1, column=0, columnspan=3, pady=(0, 6))

    # 原来放在顶部的“提示” Label 删除，不再使用单独标签

    ttk.Label(root, text="工作目录:").grid(row=2, column=0, sticky="e", padx=8, pady=8)
    path_var = tk.StringVar(value=str(Path.home()))
    ttk.Entry(root, textvariable=path_var, width=60).grid(row=2, column=1, sticky="we", padx=5, pady=8)
    ttk.Button(root, text="浏览...", command=lambda: path_var.set(filedialog.askdirectory(title="选择工作目录") or path_var.get())).grid(row=2, column=2, padx=5, pady=8)

    # 输出栏
    log_box = tk.Text(root, width=96, height=18, state="disabled")
    log_box.grid(row=4, column=0, columnspan=3, padx=10, pady=(5,10), sticky="nsew")
    root.columnconfigure(1, weight=1); root.rowconfigure(4, weight=1)

    # ★ 初始化时将提示写入输出栏
    tip = (
        "提示：若要新增通讯录，请在工作目录下放置文件名中包含“通讯录.xlsx”的文件（例如：李四-通讯录.xlsx），并至少包含以下列：姓名、实际工作单位、号码。"
    )
    log_box.config(state="normal")
    log_box.insert("end", tip + "\n")
    log_box.config(state="disabled")

    def log(msg):
        log_box.config(state="normal")
        log_box.insert("end", f"{datetime.datetime.now():%H:%M:%S}  {msg}\n")
        log_box.config(state="disabled"); log_box.see("end")

    def run(path):
        # ★ 点击开始分析后，先清空输出栏（删除提示）
        log_box.config(state="normal")
        log_box.delete("1.0", "end")
        log_box.config(state="disabled")

        global OUT_DIR
        OUT_DIR = Path(path).expanduser().resolve() / "批量分析结果"
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        _orig_print = builtins.print
        builtins.print = lambda *a, **k: log(" ".join(map(str, a)))
        try:
            all_txn = merge_all_txn(path)
            if all_txn.empty:
                messagebox.showinfo("完成", "未找到可分析文件"); return
            for person, df_person in all_txn.groupby("查询对象", dropna=False):
                print(f"--- 分析 {person} ---")
                analysis_txn(df_person)
                make_partner_summary(df_person)
            messagebox.showinfo("完成", f"全部分析完成！结果在:\n{OUT_DIR}")
        except Exception as e:
            messagebox.showerror("错误", str(e))
        finally:
            builtins.print = _orig_print

    ttk.Button(root, text="开始分析", command=lambda: threading.Thread(target=run, args=(path_var.get().strip(),), daemon=True).start(), width=18).grid(row=3, column=1, pady=10)
    root.mainloop()


# ------------------------------------------------------------------
if __name__ == "__main__":
    create_gui()
