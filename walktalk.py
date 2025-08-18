import re
import datetime
import pathlib
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from zipfile import BadZipFile
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import sys                                 # 打包兼容

# ================== 0. 路径 & 常量 ==================
def resource_path(rel: str | pathlib.Path) -> pathlib.Path:
    """打包前/后均可用的资源绝对路径"""
    if getattr(sys, "frozen", False):           # EXE 中
        base = pathlib.Path(sys._MEIPASS)
    else:                                       # 源代码
        base = pathlib.Path(__file__).resolve().parent
    return base / rel

# ➤ exe/脚本所在目录
if getattr(sys, "frozen", False):
    APP_DIR = pathlib.Path(sys.executable).resolve().parent
else:
    APP_DIR = pathlib.Path(__file__).resolve().parent

# ➤ 所有输出集中到这里
OUTPUT_DIR = APP_DIR / "输出文件"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_DIR = resource_path("走读式模板V2.2")

PLACEHOLDER_RE = re.compile(r"`([^`]+)`")

SPECIAL_TEMPLATE_STEM = "A02办案安全承诺书"
ROLE_PHRASE = "(办案组组长、办案人员、安全员)"
ROLES = ["办案组组长", "办案人员", "安全员"]

JIJIAN_OFFICES = [f"第{c}纪检监察室" for c in "一二三四五六"]
GENDER_OPTIONS = ["男", "女"]
RISK_OPTIONS   = ["低", "高"]

GROUP_DEFS = {
    "基本信息": ["第_纪检监察室", "填表日期"],
    "对象信息": ["对象姓名", "对象职务", "对象性别", "对象身份证号码", "谈话风险"],
    "时间地点": ["安全首课时间", "安全预案时间", "谈话方案时间",
               "谈话时间", "回访时间", "安全首课地点", "谈话地点"],
    "纪委人员": ["直接责任人", "第一责任人", "首谈领导",
               "主谈人", "参与人1","参与人2", "安全员", "安全员电话", "批准自行往返领导"],
}

# ================== 1. Word 替换工具 ==================
def iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

def find_placeholders(doc):
    ph = set()
    for p in iter_paragraphs(doc):
        ph.update(PLACEHOLDER_RE.findall("".join(r.text for r in p.runs)))
    return ph

def replace_in_paragraph(p, mapping):
    txt_org = "".join(r.text for r in p.runs)
    txt = txt_org
    for k, v in mapping.items():
        txt = txt.replace(f"`{k}`", str(v))
    if txt != txt_org:
        runs = p.runs or [p.add_run("")]
        runs[0].text = txt
        for r in runs[1:]:
            r.text = ""

def replace_everywhere(doc, mapping):
    for p in iter_paragraphs(doc):
        replace_in_paragraph(p, mapping)

def replace_role_phrase(doc, role):
    for p in iter_paragraphs(doc):
        txt_org = "".join(r.text for r in p.runs)
        txt = txt_org.replace(ROLE_PHRASE, role)
        if txt != txt_org:
            runs = p.runs or [p.add_run("")]
            runs[0].text = txt
            for r in runs[1:]:
                r.text = ""

def normalize_date(text):
    if not isinstance(text, str):
        try:
            text = text.strftime("%Y-%m-%d")
        except Exception:
            return str(text)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            d = datetime.datetime.strptime(text.strip(), fmt).date()
            return f"{d.year}年{d.month:02d}月{d.day:02d}日"
        except ValueError:
            continue
    return text

# ★ 新增：参与人2 统一格式化（空/“无”=>空串；否则前加“、”）
def format_canyuren2(val):
    """参与人2：空/无 => 空串；否则前面补一个 '、'（若已有则不重复）"""
    s = (val or "").strip()
    s_plain = re.sub(r"[。\s]", "", s)
    if s_plain == "" or s_plain == "无":
        return ""
    return s if s.startswith("、") else ("、" + s)

# ================== 2. Excel 记录工具 ==================
def adjust_column_widths(xls_path):
    wb = openpyxl.load_workbook(xls_path)
    ws = wb.active
    for col in ws.columns:
        first_cell = next(iter(col))
        letter = get_column_letter(first_cell.column)
        width  = max(len(str(c.value)) if c.value else 0 for c in col) + 2
        ws.column_dimensions[letter].width = width
    wb.save(xls_path)
    wb.close()  # ★ 关键：显式关闭，避免被系统认为仍在占用

def append_to_office_excel(record):
    office = record.get("第_纪检监察室", "").strip() or "未指定科室"
    path   = OUTPUT_DIR / f"{office}.xlsx"
    df_new = pd.DataFrame([record])

    if path.exists():
        df_old = pd.read_excel(path, dtype=str, engine="openpyxl")
        # 两边列对齐
        for col in df_new.columns.difference(df_old.columns):
            df_old[col] = ""
        for col in df_old.columns.difference(df_new.columns):
            df_new[col] = ""
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new

    # ★ 原子写入，避免文件被占用或中途损坏
    tmp = path.with_suffix(".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False)

    if path.exists():
        try:
            path.unlink()
        except PermissionError:
            # 若被占用，备份旧文件后再覆盖
            try:
                path.rename(path.with_suffix(".bak.xlsx"))
            except Exception:
                pass
    tmp.rename(path)

    adjust_column_widths(path)

# ================== 3. 生成 Word 文档 ==================
def generate_docs(mapping):
    obj   = mapping.get("对象姓名", "未命名") or "未命名"
    out   = OUTPUT_DIR / f"{obj}-走读式谈话"         # ➤ 输出到统一目录
    out.mkdir(parents=True, exist_ok=True)
    files = []

    for tpl in good_templates:
        if tpl.stem == SPECIAL_TEMPLATE_STEM:
            for role in ROLES:
                doc = Document(tpl)
                replace_everywhere(doc, mapping)
                replace_role_phrase(doc, role)
                name = tpl.stem
                for k, v in mapping.items():
                    name = name.replace(f"`{k}`", str(v))
                path = out / f"{name}-{role}.docx"
                doc.save(path); files.append(path)
        else:
            doc = Document(tpl)
            replace_everywhere(doc, mapping)
            name = tpl.stem
            for k, v in mapping.items():
                name = name.replace(f"`{k}`", str(v))
            path = out / f"{name}.docx"
            doc.save(path); files.append(path)
    return files

def process_record(mapping, collector: set, write_excel=True):
    # 统一处理日期/时间
    for k in mapping:
        if ("日期" in k) or ("时间" in k):
            mapping[k] = normalize_date(mapping[k])

    # ★ 统一处理“参与人2”
    if "参与人2" in mapping:
        mapping["参与人2"] = format_canyuren2(mapping.get("参与人2"))

    if write_excel:
        append_to_office_excel(mapping)

    collector.update(generate_docs(mapping))       # set 去重

# ================== 4. 扫描模板 & 占位符 ==================
good_templates, bad = [], []
for p in TEMPLATE_DIR.rglob("*.docx"):
    try:
        Document(p)
        good_templates.append(p)
    except (BadZipFile, PackageNotFoundError):
        bad.append(p)
if not good_templates:
    messagebox.showerror("未找到模板", "未发现可解析的 .docx 文件。")
    raise SystemExit

all_ph = set()
for tpl in good_templates:
    all_ph.update(find_placeholders(Document(tpl)))

defaults = {"填表日期": datetime.date.today().strftime("%Y-%m-%d")}
group_keys = {g: [] for g in GROUP_DEFS}; group_keys["其他字段"] = []
for k in all_ph:
    for g, lst in GROUP_DEFS.items():
        if k in lst:
            group_keys[g].append(k)
            break
    else:
        group_keys["其他字段"].append(k)

# ================== 5. 向导填写 ==================
def wizard_fill():
    mapping, page_vars, pages, page_keys = {}, {}, [], []
    wiz = tk.Toplevel()
    wiz.title("填写走读式谈话信息")
    wiz.resizable(False, False)

    finished = False                             # 是否点了“结束并生成”

    def is_date(k): return ("时间" in k) or ("日期" in k)

    # -------- 生成各分组页 --------
    for grp, keys in group_keys.items():
        if not keys:
            continue
        frm = ttk.Frame(wiz, padding=18)
        ttk.Label(frm, text=grp, font=("微软雅黑", 12, "bold")
                 ).grid(columnspan=3, pady=(0, 10))
        for r, key in enumerate(sorted(keys), 1):
            ttk.Label(frm, text=key + "：").grid(row=r, column=0,
                                                sticky="e", padx=6, pady=4)
            # 选项型字段
            if key == "第_纪检监察室":
                var = tk.StringVar(value=JIJIAN_OFFICES[0])
                ent = ttk.Combobox(frm, width=28, state="readonly",
                                   values=JIJIAN_OFFICES, textvariable=var)
            elif key == "对象性别":
                var = tk.StringVar(value=GENDER_OPTIONS[0])
                ent = ttk.Combobox(frm, width=8, state="readonly",
                                   values=GENDER_OPTIONS, textvariable=var)
            elif key == "谈话风险":
                var = tk.StringVar(value=RISK_OPTIONS[0])
                ent = ttk.Combobox(frm, width=8, state="readonly",
                                   values=RISK_OPTIONS, textvariable=var)
            # ★ 特殊：参与人2 —— 如无请留空
            elif key == "参与人2":
                d = defaults.get(key, "")
                var = tk.StringVar(value=d)
                ent = ttk.Entry(frm, width=30, textvariable=var)
                ttk.Label(frm, text="（如无请留空）").grid(row=r, column=2, sticky="w", padx=2)
            # 普通文本
            else:
                d = defaults.get(key, "")
                if not d and is_date(key):
                    d = datetime.date.today().strftime("%Y-%m-%d")
                width = 60 if grp == "其他字段" else 30
                var = tk.StringVar(value=d)
                ent = ttk.Entry(frm, width=width, textvariable=var)
            ent.grid(row=r, column=1, sticky="w", padx=6, pady=4)
            page_vars[key] = var
        pages.append(frm)
        page_keys.append(keys)

    # -------- 导航区 --------
    nav = ttk.Frame(wiz, padding=(12, 6))
    nav.pack(side="bottom", fill="x")
    b_prev = ttk.Button(nav, text="← 上一步")
    b_next = ttk.Button(nav, text="下一步 →")
    b_done = ttk.Button(nav, text="结束并生成", state="disabled")
    b_prev.grid(row=0, column=0, padx=4)
    b_next.grid(row=0, column=1, padx=4)
    b_done.grid(row=0, column=2, padx=4)

    cur = 0
    def show(i):
        nonlocal cur
        pages[cur].pack_forget()
        cur = i
        pages[cur].pack()
        b_prev['state'] = 'normal' if cur > 0 else 'disabled'
        b_next['state'] = 'normal' if cur < len(pages) - 1 else 'disabled'
        b_done['state'] = 'normal' if cur == len(pages) - 1 else 'disabled'

    # ★ 放宽校验：允许“参与人2”留空
    def valid(i):
        return all(page_vars[k].get().strip() for k in page_keys[i] if k != "参与人2")

    b_prev.config(command=lambda: show(cur - 1))

    def nxt(e=None):
        if not valid(cur):
            messagebox.showwarning("请完善", "当前页还有未填写字段！", parent=wiz)
            return
        if cur < len(pages) - 1:
            show(cur + 1)
    b_next.config(command=nxt)
    wiz.bind("<Return>", nxt)

    def fin(e=None):
        nonlocal finished
        # ★ 结束前的未填提示不包含“参与人2”
        unfilled = [k for k, v in page_vars.items() if (k != "参与人2") and (not v.get().strip())]
        if unfilled and not messagebox.askyesno(
                "尚有未填字段",
                "以下字段为空：\n" + "\n".join(unfilled) + "\n\n仍要生成文档吗？",
                parent=wiz):
            return
        finished = True
        wiz.destroy()
    b_done.config(command=fin)

    # “×” 或 Esc 直接关闭，不生成
    wiz.bind("<Escape>", lambda e: wiz.destroy())
    wiz.protocol("WM_DELETE_WINDOW", wiz.destroy)

    pages[0].pack()
    b_prev["state"] = "disabled"
    wiz.grab_set()
    wiz.wait_window()

    if not finished:
        return None                      # 未完成，直接返回
    mapping.update({k: v.get().strip() for k, v in page_vars.items()})
    return mapping

# ================== 6. Excel 导入 ==================
def import_from_excel():
    f = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
    if not f:
        return
    try:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("读取失败", f"无法读取 Excel：\n{e}", parent=menu)
        return
    if "对象姓名" not in df.columns:
        messagebox.showwarning("列缺失", "Excel 必须包含 “对象姓名” 列！", parent=menu)
        return
    df.fillna("", inplace=True)

    sel = tk.Toplevel()
    sel.title("选择生成对象")
    sel.resizable(False, False)
    ttk.Label(sel, text="勾选需要生成的对象：",
              font=("微软雅黑", 10, "bold")).pack(pady=(10, 6))
    lb = tk.Listbox(sel, selectmode="extended", width=40, height=15)
    for i, n in enumerate(df["对象姓名"]):
        lb.insert(i, n)
    lb.pack(padx=12)
    frm = ttk.Frame(sel, padding=6)
    frm.pack()
    ttk.Button(frm, text="全选",
               command=lambda: lb.selection_set(0, tk.END)).grid(row=0, column=0, padx=4)
    gen = ttk.Button(frm, text="生成")
    gen.grid(row=0, column=1, padx=4)

    outs: set[pathlib.Path] = set()
    def doit():
        idxs = lb.curselection()
        if not idxs:
            messagebox.showinfo("未选择", "请至少选择一个对象！", parent=sel)
            return
        sel.destroy()
        for i in idxs:
            rec = df.iloc[i].to_dict()
            for k in all_ph:
                rec.setdefault(k, "")
            process_record(rec, outs, write_excel=False)  # ★ 处理时会自动规范参与人2
        paths = sorted(map(str, outs))
        messagebox.showinfo("完成", "已生成以下文件：\n" + "\n".join(paths), parent=menu)
    gen.config(command=doit)
    sel.grab_set()

# ================== 7. 主菜单 ==================
root = tk.Tk()
root.withdraw()
menu = tk.Toplevel()
menu.title("温岭纪委走读式谈话文件生成工具")
menu.geometry("500x300")
menu.resizable(False, False)

ttk.Label(menu, text="温岭纪委走读式谈话文件生成工具",
          font=("微软雅黑", 16, "bold")).pack(pady=(20, 4))
ttk.Label(menu, text="© 温岭纪委六室 单柳昊",
          font=("微软雅黑", 9)).pack()

btn_frame = ttk.Frame(menu, padding=20)
btn_frame.pack(expand=True)
b_manual = ttk.Button(btn_frame, text="手动填写", width=20)
b_excel  = ttk.Button(btn_frame, text="从 Excel 导入", width=20)
b_manual.grid(row=0, column=0, pady=6)
b_excel.grid(row=1, column=0, pady=6)
ttk.Label(menu, text="生成结果仅供参考，请仔细检查",
          font=("微软雅黑", 8)).place(relx=1.0, rely=1.0,
                                      anchor="se", x=-6, y=-4)

def run_manual():
    m = wizard_fill()
    if m:
        outs: set[pathlib.Path] = set()
        process_record(m, outs)
        paths = sorted(map(str, outs))
        messagebox.showinfo("完成", "已生成以下文件：\n" + "\n".join(paths), parent=menu)

b_manual.config(command=run_manual)
b_excel.config(command=import_from_excel)

# ★ 新增：统一退出，确保不留后台进程
def exit_app(*_):
    try:
        for w in list(root.winfo_children()):
            try:
                w.destroy()
            except Exception:
                pass
        root.quit()
    finally:
        try:
            root.destroy()
        finally:
            sys.exit(0)

menu.bind("<Escape>", exit_app)
menu.protocol("WM_DELETE_WINDOW", exit_app)

menu.mainloop()
