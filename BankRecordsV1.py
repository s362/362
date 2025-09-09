#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
交易流水批量分析工具 GUI   v6-plus (refactor + 线下银行扩展)
Author  : 温岭纪委六室 单柳昊   （2025-08-05 修订）
重构者  : （效率优化版 2025-08-28）
扩展者  : （线下农行/建行接入 2025-09-09）

（2025-08-27 增补保持不变）
- 新增：支持读取同目录下固定文件名 “交易明细信息.csv”
- “查询对象”自动来自同目录 “人员信息.csv” 的“客户姓名”
- “反馈单位”来自 CSV 文件父文件夹名

（2025-08-28 重构亮点：功能等价、性能更佳）
- 缓存 header 探测/特殊表头探测，避免多次读取同一文件首行
- 解析泰隆多 sheet 时只计算一次 header 行
- 星期/节假日基于已算好的时间戳一次性向量化生成（错误/NaT 标注保持一致）
- 尽量减少 DataFrame copy 与重复类型转换
- I/O 小优化（自动列宽逻辑保留，分支更精简）

（2025-08-28+ 职务增强）
- 自动读取目录内包含“通讯录”的Excel（多sheet），提取 姓名、主管部门名称、行政职务
- “职务”= 主管部门名称-行政职务（缺一取一）
- 与交易中的“交易对方姓名”匹配，输出“对方职务”到：
  1) 所有人-合并交易流水.xlsx
  2) 资金来源分析
  3) 交易对手分析 / 与公司相关交易频次分析

（2025-09-09 线下银行扩展）
- 新增接入：农业银行线下（识别 APSH sheet）
- 新增接入：建设银行线下（识别 “交易明细” sheet）

（2025-09-09+ 数据质量增强）
- 三键去重：若【交易流水号 + 交易时间 + 交易金额】完全一致，自动去重
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading, warnings, builtins, datetime, re
from pathlib import Path
from functools import wraps, lru_cache
from typing import Optional, List, Dict, Any

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

# 新增：全局通讯录“姓名->职务”映射（供合并及分析阶段使用）
CONTACT_TITLE_MAP: Dict[str, str] = {}

# ------------------------------------------------------------------
# ②   通用工具
# ------------------------------------------------------------------
SKIP_HEADER_KEYWORDS = [
    "反洗钱-电子账户交易明细",
    "信用卡消费明细",
]

@lru_cache(maxsize=None)
def _should_skip_special_cached(path_str: str) -> Optional[str]:
    """首 3 行包含关键字则返回关键字，否则 None；带缓存"""
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
    if OUT_DIR is not None:
        filename = OUT_DIR / filename
    filename = Path(filename).with_suffix(".xlsx")
    filename.parent.mkdir(parents=True, exist_ok=True)

    df = df.replace(np.nan, "")
    if engine == "xlsxwriter":
        with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=index)
            ws = writer.sheets[sheet_name]
            # 计算列宽（一次 map + max）
            for i, col in enumerate(df.columns):
                s = df[col].astype(str)
                width = max(
                    min_width,
                    min(max(s.map(len).max(), len(str(col))) + 2, max_width),
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
    last_err: Optional[Exception] = None
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
    vals = df.astype(str).replace("nan", "", regex=False).to_numpy().ravel().tolist()
    for val in vals:
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
            header_idx = _header_row(fp)          # 自动定位表头行（缓存后只算一次）
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
@lru_cache(maxsize=None)
def _header_row(path: Path) -> int:
    """读取文件首 15 行寻找包含“交易日期”的行号；带缓存"""
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

    try:
        df = raw.copy()
        df.columns = pd.Index(df.columns).astype(str).str.strip()
        n = len(df)

        def _S(default=""):
            return pd.Series([default] * n, index=df.index)

        def _safe(name, fn):
            try:
                return fn()
            except Exception as e:
                print(f"⚠️ CSV字段[{name}]解析异常：{e}")
                return _S("wrong")

        def col(keys, default=""):
            if isinstance(keys, str):
                return df[keys] if keys in df else _S(default)
            for k in keys:
                if k in df:
                    return df[k]
            return _S(default)

        def _to_str_no_sci(x):
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

        out = pd.DataFrame(index=df.index)

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
            {"人民币":"CNY","人民币元":"CNY","RMB":"CNY","156":"CNY"}).fillna("CNY"))

        # ===== 金额 / 余额 =====
        out["交易金额"] = _safe("交易金额", lambda: pd.to_numeric(col(["交易金额","金额","发生额"], 0), errors="coerce"))
        out["账户余额"] = _safe("账户余额", lambda: pd.to_numeric(col(["交易余额","余额","账户余额"], 0), errors="coerce"))

        # ===== 借贷/收付 =====
        out["借贷标志"] = col(["收付标志",""], "")

        # ===== 交易时间 =====
        if "交易时间" in df.columns:
            out["交易时间"] = _safe("交易时间", lambda: np.where(
                pd.to_datetime(df["交易时间"], errors="coerce").notna(),
                pd.to_datetime(df["交易时间"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M:%S"),
                df["交易时间"].astype(str)
            ))
        else:
            out["交易时间"] = _S("wrong")

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

        # ===== 备注合并 =====
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

        # 对齐模板列顺序
        out = out.reindex(columns=TEMPLATE_COLS, fill_value="")
        return out

    except Exception as e:
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
    2025-08-06  增强版：支持 dict[sheet, df] 合并；保持原行为
    """
    if isinstance(raw, dict):
        frames: List[pd.DataFrame] = []
        for sheet_name, df_sheet in raw.items():
            one = tl_to_template(df_sheet)
            if not one.empty:
                one.insert(0, "__sheet__", sheet_name)
                frames.append(one)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=TEMPLATE_COLS)

    df: pd.DataFrame = raw
    if df.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    def col_multi(keys, default=""):
        for k in keys:
            if k in df:
                return df[k]
        return pd.Series([default] * len(df), index=df.index)

    out = pd.DataFrame(index=df.index)
    out["本方账号"] = col_multi(["客户账号","账号","本方账号"], "wrong")
    out["查询账户"] = out["本方账号"]
    out["反馈单位"] = "泰隆银行"
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

    out = out.reindex(columns=TEMPLATE_COLS, fill_value="")
    return out

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

    out = pd.DataFrame(index=df.index)
    acct = col("账号卡号").astype(str).str.replace(r"\.0$", "", regex=True)
    out["本方账号"] = acct
    out["查询账户"] = acct
    out["查询对象"] = holder
    out["反馈单位"] = "民泰银行"
    out["币种"] = col("币种").astype(str).replace("人民币","CNY").replace("人民币元","CNY").fillna("CNY")

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

    out = out.reindex(columns=TEMPLATE_COLS, fill_value="")
    return out

# ------------------------------------------------------------------
# ⑤   农商行 → 模板
# ------------------------------------------------------------------
def rc_to_template(raw: pd.DataFrame, holder: str, is_old: bool) -> pd.DataFrame:
    if raw.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    def col(c, default=""):
        return raw[c] if c in raw else pd.Series([default] * len(raw), index=raw.index)

    out = pd.DataFrame(index=raw.index)
    out["本方账号"] = col("账号", "wrong")
    out["查询账户"] = out["本方账号"]
    out["交易金额"] = col("发生额") if is_old else col("交易金额")
    out["账户余额"] = col("余额") if is_old else col("交易余额")
    out["反馈单位"] = "老农商银行" if is_old else "新农商银行"

    dates = col("交易日期").astype(str)
    times = col("交易时间").astype(str)
    out["交易时间"] = [_parse_dt(d, t, is_old) for d, t in zip(dates, times)]

    out["借贷标志"] = col("借贷标志")
    out["币种"] = "CNY" if is_old else col("币种").replace("人民币","CNY").replace("人民币元","CNY")
    out["查询对象"] = holder
    out["交易对方姓名"] = col("对方姓名", " ")
    out["交易对方账户"] = col("对方账号", " ")
    out["交易网点名称"] = col("代理行机构号") if is_old else col("交易机构")
    out["交易摘要"] = col("备注") if is_old else col("摘要", "wrong")

    out = out.reindex(columns=TEMPLATE_COLS, fill_value="")
    return out

# ===============================
# ⑤.8  农业银行线下（APSH） → 模板（合并 yyyymmdd + HHMMSS）
# ===============================
def _is_abc_offline_file(p: Path) -> bool:
    """是否为农行线下查询格式：含 APSH sheet。"""
    try:
        xls = pd.ExcelFile(p)
        return "APSH" in xls.sheet_names
    except Exception:
        return False

def _merge_abc_datetime(date_val, time_val) -> str:
    """
    将 yyyymmdd 与 时间(无连接符 HHMMSS，或 13:31:20，或 Excel 小数时间，或空) 合并为 'YYYY-MM-DD HH:MM:SS'。
    规则：
      - 交易时间为空/NaN/空字符串 => 00:00:00
      - 交易时间为 Excel 小数(0~1) => 按一天的秒数换算
      - 纯数字长度<6 左补零，>6 取前 6 位
      - 示例：20100113 + 133120 -> 2010-01-13 13:31:20
    """
    # ---- 日期处理 ----
    ds_raw = "" if date_val is None else str(date_val).strip()
    ds_digits = re.sub(r"\D", "", ds_raw)
    date_ts = None
    if len(ds_digits) >= 8:
        ds8 = ds_digits[:8]
        date_ts = pd.to_datetime(ds8, format="%Y%m%d", errors="coerce")
    else:
        # 兜底：直接尝试解析
        date_ts = pd.to_datetime(date_val, errors="coerce")
    if pd.isna(date_ts):
        return "wrong"
    date_str = date_ts.strftime("%Y-%m-%d")

    # ---- 时间处理：统一得到 'HHMMSS' 的 6 位字符串 ----
    def to_hhmmss_str(t) -> str:
        # 空、NaN、None -> 00:00:00
        if t is None or (isinstance(t, float) and np.isnan(t)) or (isinstance(t, str) and t.strip() == "") or pd.isna(t):
            return "000000"

        # Excel 小数时间（0~1）
        if isinstance(t, (int, np.integer)) or isinstance(t, (float, np.floating)):
            try:
                tf = float(t)
                if 0.0 <= tf < 1.0:
                    secs = int(round(tf * 86400))
                    if secs >= 86400:
                        secs = 0  # 极端四舍五入到 24:00:00，当作 00:00:00
                    h = secs // 3600
                    m = (secs % 3600) // 60
                    s = secs % 60
                    return f"{h:02d}{m:02d}{s:02d}"
                # 常见：133120.0 / 93120.0
                digits = re.sub(r"\D", "", str(int(round(tf))))
                if len(digits) < 6:
                    digits = digits.zfill(6)
                else:
                    digits = digits[:6]
                return digits
            except Exception:
                pass

        # 字符串：可能是 '13:31:20' / '13.31.20' / '133120' / '93120'
        s = str(t).strip()
        # 带分隔符的情况，尝试按时间解析
        if ":" in s or "." in s:
            s2 = s.replace(".", ":")
            tt = pd.to_datetime("2000-01-01 " + s2, errors="coerce")
            if pd.notna(tt):
                return tt.strftime("%H%M%S")
        # 纯提取数字
        digits = re.sub(r"\D", "", s)
        if digits == "":
            return "000000"
        if len(digits) < 6:
            digits = digits.zfill(6)
        else:
            digits = digits[:6]
        return digits

    hhmmss = to_hhmmss_str(time_val)
    hh, mm, ss = hhmmss[:2], hhmmss[2:4], hhmmss[4:6]
    return f"{date_str} {hh}:{mm}:{ss}"

def abc_offline_from_file(p: Path) -> pd.DataFrame:
    """
    农业银行线下查询（APSH）流水 → 统一模板字段 TEMPLATE_COLS
    适配列（常见）：账号、交易日期(yyyymmdd)、交易时间(HHMMSS，无连接符)、卡号、户名、传票号、交易网点、交易金额、交易后余额、
             摘要、交易渠道、对方账号、对方户名、对方开户行、交易行号
    —— 本函数将【交易日期 + 交易时间】合并生成标准“交易时间(YYYY-MM-DD HH:MM:SS)”。
    """
    try:
        xls = pd.ExcelFile(p)
        if "APSH" not in xls.sheet_names:
            return pd.DataFrame(columns=TEMPLATE_COLS)
        df = xls.parse("APSH", header=0)
    except Exception:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    if df.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    # 列名清洗
    df.columns = pd.Index(df.columns).astype(str).str.strip()
    n = len(df)
    out = pd.DataFrame(index=df.index)

    # 本方/查询账号卡号
    out["本方账号"] = df.get("账号", "")
    out["本方卡号"] = df.get("卡号", "").astype(str).str.replace(r"\.0$", "", regex=True)
    out["查询账户"] = out["本方账号"]
    out["查询卡号"] = out["本方卡号"]

    # 查询对象/反馈单位/币种
    holder = df.get("户名", "")
    if not isinstance(holder, pd.Series):
        holder = pd.Series([holder]*n, index=df.index)
    out["查询对象"] = holder.fillna("").astype(str).str.strip().replace({"nan": ""}).replace("", "未知")
    out["反馈单位"] = "农业银行"
    out["币种"] = "CNY"

    # 金额/余额/借贷标志（按正负号判断）
    amt = pd.to_numeric(df.get("交易金额", 0), errors="coerce")
    out["交易金额"] = amt
    out["账户余额"] = pd.to_numeric(df.get("交易后余额", ""), errors="coerce")
    out["借贷标志"] = np.where(amt > 0, "进", np.where(amt < 0, "出", ""))

    # === 交易时间：合并 yyyymmdd + HHMMSS（无连接符） ===
    dates = df.get("交易日期", "")
    times = df.get("交易时间", "")
    out["交易时间"] = [_merge_abc_datetime(d, t) for d, t in zip(dates, times)]

    # 其它字段对齐
    out["交易摘要"] = df.get("摘要", "").astype(str)
    out["交易流水号"] = ""  # APSH 多无此字段
    out["交易类型"] = ""    # 可根据需要由 摘要/渠道 推断；此处留空
    out["交易对方姓名"] = df.get("对方户名", " ").astype(str)
    out["交易对方账户"] = df.get("对方账号", " ").astype(str)
    out["交易对方卡号"] = ""
    out["交易对方证件号码"] = " "
    out["交易对手余额"] = ""
    out["交易对方账号开户行"] = df.get("对方开户行", " ").astype(str)
    out["交易网点名称"] = df.get("交易网点", "").astype(str)
    out["交易网点代码"] = df.get("交易行号", "").astype(str)
    out["日志号"] = ""
    out["传票号"] = df.get("传票号", "").astype(str)
    out["凭证种类"] = ""
    out["凭证号"] = ""
    out["现金标志"] = ""
    out["终端号"] = df.get("交易渠道", "").astype(str)
    out["交易是否成功"] = ""
    out["交易发生地"] = ""
    out["商户名称"] = ""
    out["商户号"] = ""
    out["IP地址"] = ""
    out["MAC"] = ""
    out["交易柜员号"] = ""
    out["备注"] = ""

    # 模板列顺序
    out = out.reindex(columns=TEMPLATE_COLS, fill_value="")
    return out

# ===============================
# ⑤.9  建设银行线下（交易明细） → 模板（新增）
# ===============================
def _is_ccb_offline_file(p: Path) -> bool:
    """
    粗识别建设银行线下：存在名为“交易明细”的sheet，且包含关键字段。
    """
    try:
        xls = pd.ExcelFile(p)
        if "交易明细" not in xls.sheet_names:
            return False
        # 取头一行看列名是否含关键字段
        df_head = xls.parse("交易明细", nrows=1)
        cols = set(map(str, df_head.columns))
        required = {"客户名称", "账号", "交易日期", "交易时间", "交易金额"}
        return required.issubset(cols)
    except Exception:
        return False

def ccb_offline_from_file(p: Path) -> pd.DataFrame:
    """
    建设银行线下（交易明细） → 统一模板字段
    适配列：客户名称、账号、交易日期、交易时间、交易卡号、摘要、借贷方向、交易金额、账户余额、
          柜员号、交易机构号、交易机构名称、对方账号、对方户名、对方行名、交易流水号、交易渠道、
          自助设备编号、扩充备注、IP地址、MAC地址、第三方订单号、商户号、商户名称
    """
    try:
        xls = pd.ExcelFile(p)
        if "交易明细" not in xls.sheet_names:
            return pd.DataFrame(columns=TEMPLATE_COLS)
        df = xls.parse("交易明细", header=0)
    except Exception:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    if df.empty:
        return pd.DataFrame(columns=TEMPLATE_COLS)

    df.columns = pd.Index(df.columns).astype(str).str.strip()
    out = pd.DataFrame(index=df.index)

    # 基本字段
    out["本方账号"] = df.get("账号", "")
    out["本方卡号"] = df.get("交易卡号", "").astype(str).str.replace(r"\.0$", "", regex=True)
    out["查询账户"] = out["本方账号"]
    out["查询卡号"] = out["本方卡号"]

    out["查询对象"] = df.get("客户名称", "").astype(str).replace({"nan":""}).replace("", "未知")
    out["反馈单位"] = "建设银行"
    out["币种"] = df.get("币种", "CNY").astype(str).replace({"人民币":"CNY","人民币元":"CNY","RMB":"CNY","156":"CNY"}).fillna("CNY")

    amt = pd.to_numeric(df.get("交易金额", 0), errors="coerce")
    out["交易金额"] = amt
    out["账户余额"] = pd.to_numeric(df.get("账户余额", ""), errors="coerce")

    # 借贷方向：借->出，贷->进
    jd = df.get("借贷方向", "").astype(str).str.strip()
    out["借贷标志"] = np.where(jd.str.contains("^贷", na=False) | jd.str.upper().isin(["贷","C","CR","CREDIT"]), "进",
                        np.where(jd.str.contains("^借", na=False) | jd.str.upper().isin(["借","D","DR","DEBIT"]), "出",
                                 np.where(amt>0, "进", np.where(amt<0, "出", ""))))

    # 时间
    dates = df.get("交易日期", "")
    times = df.get("交易时间", "")
    times_str = pd.Series(times).astype(str).str.replace(r"\.0$", "", regex=True)
    out["交易时间"] = [_parse_dt(d, t, is_old=False) for d, t in zip(dates, times_str)]

    # 其它映射
    out["交易摘要"] = df.get("摘要", " ").astype(str)
    out["交易类型"] = ""  # 保留空位（如需由摘要/渠道二次推断可自行扩展）
    out["交易流水号"] = df.get("交易流水号", "").astype(str)

    out["交易对方姓名"] = df.get("对方户名", " ").astype(str)
    out["交易对方账户"] = df.get("对方账号", " ").astype(str)
    out["交易对方卡号"] = ""
    out["交易对方证件号码"] = " "
    out["交易对手余额"] = ""
    out["交易对方账号开户行"] = df.get("对方行名", " ").astype(str)

    out["交易网点名称"] = df.get("交易机构名称", "").astype(str)
    out["交易网点代码"] = df.get("交易机构号", "").astype(str)
    out["交易柜员号"] = df.get("柜员号", "").astype(str)

    out["终端号"] = df.get("交易渠道", "").astype(str)  # 常见形态：渠道代码
    # 其它可用补充信息 → 备注
    ext = df.get("扩充备注", "").astype(str).replace({"nan":""})
    out["备注"] = ext

    out["现金标志"] = ""
    out["日志号"] = ""
    out["传票号"] = ""
    out["凭证种类"] = ""
    out["凭证号"] = ""

    out["交易是否成功"] = ""
    out["交易发生地"] = ""

    out["商户名称"] = df.get("商户名称", "").astype(str)
    out["商户号"] = df.get("商户号", "").astype(str)
    out["IP地址"] = df.get("IP地址", "").astype(str)
    out["MAC"] = df.get("MAC地址", "").astype(str)

    # 对齐模板
    out = out.reindex(columns=TEMPLATE_COLS, fill_value="")
    return out

# ------------------------------------------------------------------
# ⑤.5  通讯录读取与职务匹配（新增）
# ------------------------------------------------------------------
CONTACT_NAME_COLS = ["姓名", "联系人", "人员姓名", "姓名/名称"]
DEPT_COLS = ["主管部门名称", "部门", "所属部门", "归属单位", "单位", "工作单位", "科室", "处室", "所属单位"]
TITLE_COLS = ["行政职务", "岗位", "职称"]

def _find_first_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in df.columns:
        cs = str(c).strip()
        for key in candidates:
            if key in cs:
                return c
    return None

def _compose_title_str(dept: str, title: str) -> str:
    """职务拼接规则：部门-行政职务；缺一取一；都空则空"""
    def _blank(x: Any) -> bool:
        s = str(x).strip() if x is not None else ""
        return s == "" or s.lower() in {"nan", "none"} or s in {"-", "—", "——", "无", "暂无"}
    d = "" if _blank(dept) else str(dept).strip()
    t = "" if _blank(title) else str(title).strip()

    if d and t:
        return f"{d}-{t}"
    if t:
        return t
    if d:
        return d
    return ""

def _read_one_contacts_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    # 先尝试直接读头，再退回扫描前10行找“姓名”
    try:
        df0 = xls.parse(sheet_name=sheet_name, header=0)
        df0.columns = pd.Index(df0.columns).astype(str).str.strip()
        if any("姓名" in c for c in df0.columns):
            df = df0
        else:
            raise ValueError("未命中姓名列，尝试扫描表头")
    except Exception:
        head = xls.parse(sheet_name=sheet_name, header=None, nrows=10)
        header_idx = 0
        for i, row in head.iterrows():
            if row.astype(str).str.contains("姓名").any():
                header_idx = i
                break
        df = xls.parse(sheet_name=sheet_name, header=header_idx)
        df.columns = pd.Index(df.columns).astype(str).str.strip()

    name_col  = _find_first_col(df, CONTACT_NAME_COLS)
    dept_col  = _find_first_col(df, DEPT_COLS)
    title_col = _find_first_col(df, TITLE_COLS)
    if not name_col:
        return pd.DataFrame(columns=["姓名","职务"])

    out = pd.DataFrame()
    out["姓名"] = df[name_col].astype(str).str.strip()
    dept = df[dept_col].astype(str) if dept_col in df else ""
    titl = df[title_col].astype(str) if title_col in df else ""
    if isinstance(dept, str): dept = pd.Series([dept]*len(out))
    if isinstance(titl, str): titl = pd.Series([titl]*len(out))
    out["职务"] = [_compose_title_str(d, t) for d, t in zip(dept, titl)]
    out = out[(out["姓名"]!="") & (~out["姓名"].str.lower().eq("nan"))]
    out["职务"] = out["职务"].replace({"nan":"","None":""})
    out.drop_duplicates(inplace=True)
    return out[["姓名","职务"]]

def load_contacts_map(root: Path) -> Dict[str, str]:
    files = [p for p in root.rglob("*通讯录*.xls*")]
    if not files:
        print("ℹ️ 未发现文件名包含“通讯录”的Excel。")
        return {}

    frames: List[pd.DataFrame] = []
    for p in files:
        try:
            xls = pd.ExcelFile(p)
        except Exception as e:
            print("❌ 通讯录载入失败", p.name, e)
            continue
        for sht in xls.sheet_names:
            try:
                df = _read_one_contacts_sheet(xls, sht)
                if not df.empty:
                    df["来源文件"] = p.name
                    df["来源sheet"] = sht
                    frames.append(df)
            except Exception as e:
                print("❌ 通讯录解析失败", f"{p.name}->{sht}", e)

    if not frames:
        print("ℹ️ 通讯录中未解析出有效姓名/职务。")
        return {}

    allc = pd.concat(frames, ignore_index=True)
    allc["姓名"] = allc["姓名"].astype(str).str.strip()
    allc["职务"] = allc["职务"].astype(str).str.strip()

    def _uniq_preserve(seq: List[str]) -> List[str]:
        """去空去'nan'并按出现顺序去重"""
        seen = set()
        out: List[str] = []
        for x in seq:
            s = (x or "").strip()
            if not s or s.lower() == "nan":
                continue
            if s not in seen:
                seen.add(s)
                out.append(s)
        return out

    # 同名多条 -> 合并成 '、' 分隔的一条
    grouped = (allc.groupby("姓名")["职务"]
                    .apply(lambda s: "、".join(_uniq_preserve(list(s)))))
    mapping = grouped.to_dict()

    print(f"✅ 已读取通讯录 {len(files)} 份，收录姓名 {len(mapping)} 条（同名已合并）。")
    return mapping

# ------------------------------------------------------------------
# ⑥   合并全部流水
# ------------------------------------------------------------------
def merge_all_txn(root_dir: str) -> pd.DataFrame:
    root = Path(root_dir).expanduser().resolve()

    # —— 新增：先加载通讯录（全局映射）
    global CONTACT_TITLE_MAP
    CONTACT_TITLE_MAP = load_contacts_map(root)

    china_files = [p for p in root.rglob("*-*-交易流水.xls*")]

    all_excel = list(root.rglob("*.xls*"))
    rc_candidates = [p for p in all_excel if "农商行" in p.as_posix()]
    pattern_old = re.compile(r"老\s*[账帐]\s*(?:号|户)")
    old_rc = [p for p in rc_candidates if pattern_old.search(p.stem)]
    new_rc = [p for p in rc_candidates if p not in old_rc]

    tl_files = [p for p in all_excel if "泰隆" in p.as_posix()]
    mt_files = [p for p in all_excel if "民泰" in p.as_posix()]

    # —— 新增：农行线下（APSH）、建行线下（交易明细）
    abc_offline_files = [p for p in all_excel if _is_abc_offline_file(p)]
    ccb_offline_files = [p for p in all_excel if _is_ccb_offline_file(p)]

    csv_txn_files = [p for p in root.rglob("交易明细信息.csv")]

    print(
        f"✅ 网上银行 {len(china_files)} 份，"
        f"老农商 {len(old_rc)} 份，新农商 {len(new_rc)} 份，"
        f"泰隆银行 {len(tl_files)} 份，"
        f"民泰银行 {len(mt_files)} 份，"
        f"农行线下 {len(abc_offline_files)} 份，"
        f"建行线下 {len(ccb_offline_files)} 份，"
        f"交易明细CSV {len(csv_txn_files)} 份"
    )

    dfs: List[pd.DataFrame] = []

    # —— 网银标准表 —— 直接读入
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

    # —— 农商行 —— 新旧分流
    for p in old_rc + new_rc:
        print(f"正在处理 {p.name} ...")
        kw = should_skip_special(p)
        if kw:
            print(f"⏩ 跳过【{p.name}】：表头含“{kw}”")
            continue

        raw = _read_raw(p)

        holder = extract_holder_from_df(raw) or holder_from_folder(p.parent) or fallback_holder_from_path(p)
        is_old = p in old_rc
        df = rc_to_template(raw, holder, is_old)
        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)

    # —— 泰隆 —— 
    for p in tl_files:
        if "开户" in p.stem:
            continue
        print(f"正在处理 {p.name} ...")
        try:
            xls = pd.ExcelFile(p)
        except Exception as e:
            print("❌", f"{p.name} 载入失败", e)
            continue

        try:
            header_idx = _header_row(p)  # 缓存后仅计算一次
        except Exception as e:
            print("❌", f"{p.name} 表头行识别失败", e)
            header_idx = 0

        xls_dict: Dict[str, pd.DataFrame] = {}
        for sht in xls.sheet_names:
            try:
                df_sheet = xls.parse(sheet_name=sht, header=header_idx)
                xls_dict[sht] = df_sheet
            except Exception as e:
                print("❌", f"{p.name} -> {sht}", e)

        df = tl_to_template(xls_dict)
        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)

    # —— 民泰 —— 常规
    for p in mt_files:
        print(f"正在处理 {p.name} ...")
        raw = _read_raw(p)
        df  = mt_to_template(raw)
        if not df.empty:
            df["来源文件"] = p.name
            dfs.append(df)

    # —— 农行线下 —— APSH
    for p in abc_offline_files:
        print(f"正在处理 {p.name} ...")
        try:
            df = abc_offline_from_file(p)
            if not df.empty:
                df["来源文件"] = p.name
                dfs.append(df)
        except Exception as e:
            print("❌ 农行线下解析失败", p.name, e)

    # —— 建行线下 —— 交易明细
    for p in ccb_offline_files:
        print(f"正在处理 {p.name} ...")
        try:
            df = ccb_offline_from_file(p)
            if not df.empty:
                df["来源文件"] = p.name
                dfs.append(df)
        except Exception as e:
            print("❌ 建行线下解析失败", p.name, e)

    # —— 交易明细 CSV
    for p in csv_txn_files:
        print(f"正在处理 {p.name} ...")
        try:
            raw_csv = _read_csv_smart(p)
        except Exception as e:
            print("❌ 无法读取CSV", p.name, e)
            continue

        holder = _person_from_people_csv(p.parent) or holder_from_folder(p.parent) or fallback_holder_from_path(p)
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

    # —— 新增：三键去重（交易流水号 + 交易时间 + 交易金额）
    # 说明：为避免因金额格式差异导致的“假不同”，将金额先转为数值并保留两位小数再去重
    all_txn["交易金额"] = pd.to_numeric(all_txn["交易金额"], errors="coerce").round(2)
    before = len(all_txn)
    all_txn = all_txn.drop_duplicates(subset=["交易流水号", "交易时间", "交易金额"], keep="first").reset_index(drop=True)
    removed = before - len(all_txn)
    if removed:
        print(f"🧹 已按“交易流水号+交易时间+交易金额”去重 {removed} 条。")

    # —— 统一：排序、序号、类型标准化、分箱、星期/节假日 —— 向量化加速
    ts = pd.to_datetime(all_txn["交易时间"], errors="coerce")
    all_txn.insert(0, "__ts__", ts)
    all_txn.sort_values("__ts__", inplace=True, kind="mergesort")  # 稳定排序
    all_txn["序号"] = range(1, len(all_txn) + 1)
    all_txn.drop(columns="__ts__", inplace=True)

    all_txn["交易金额"] = pd.to_numeric(all_txn["交易金额"], errors="coerce")

    # 借贷标志标准化
    def _std_flag(x):
        if pd.isna(x):
            return x
        s = str(x).strip()
        if s in {"1","借","D"}: return "出"
        if s in {"2","贷","C"}: return "进"
        return s
    all_txn["借贷标志"] = all_txn["借贷标志"].apply(_std_flag)

    # 金额分箱
    bins = [-np.inf, 2000, 5000, 20000, 50000, np.inf]
    labels = ["2000以下","2000-5000","5000-20000","20000-50000","50000以上"]
    all_txn["金额区间"] = pd.cut(all_txn["交易金额"], bins=bins, labels=labels, right=False, include_lowest=True)

    # 星期（向量化）
    weekday_map = {0:"星期一",1:"星期二",2:"星期三",3:"星期四",4:"星期五",5:"星期六",6:"星期日"}
    wk = pd.Series(index=all_txn.index, dtype=object)
    mask_valid = ts.notna()
    wk.loc[mask_valid] = ts.dt.weekday.map(weekday_map)
    wk.loc[~mask_valid] = "wrong"
    all_txn["星期"] = wk

    # 节假日（对唯一日期做缓存映射）
    dates = ts.dt.date
    status = pd.Series(index=all_txn.index, dtype=object)
    unique_dates = pd.unique(dates[mask_valid])
    @lru_cache(maxsize=None)
    def _day_status(d) -> str:
        try:
            return "节假日" if is_holiday(d) else ("工作日" if is_workday(d) else "周末")
        except Exception:
            # 高稳健兜底
            dt = datetime.datetime.combine(d, datetime.time())
            return "周末" if dt.weekday() >= 5 else "工作日"
    if len(unique_dates):
        map_dict = {d: _day_status(d) for d in unique_dates}
        status.loc[mask_valid] = dates.loc[mask_valid].map(map_dict)
    status.loc[~mask_valid] = "wrong"
    all_txn["节假日"] = status

    # —— 新增：匹配通讯录，增加“对方职务”列
    if CONTACT_TITLE_MAP:
        all_txn["对方职务"] = all_txn["交易对方姓名"].map(CONTACT_TITLE_MAP).fillna("")
    else:
        all_txn["对方职务"] = ""

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
    # 新增：对方职务（来自通讯录映射）
    if CONTACT_TITLE_MAP:
        src.insert(1, "对方职务", src["交易对方姓名"].map(CONTACT_TITLE_MAP).fillna(""))

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

    # 新增：对方职务
    if CONTACT_TITLE_MAP:
        summ.insert(2, "对方职务", summ["对方姓名"].map(CONTACT_TITLE_MAP).fillna(""))

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
