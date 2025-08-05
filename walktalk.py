import re
import datetime
import pathlib
import tkinter as tk
from tkinter import simpledialog, messagebox
from docx import Document

# =================== 0. 参数 & 公用函数 ===================
TEMPLATE_DIR   = pathlib.Path(r"C:\Users\Lenovo\Desktop\workspace\走读式模板V2.2")
PLACEHOLDER_RE = re.compile(r"`([^`]+)`")

# —— 专门针对 A02办案安全承诺书 的设定 ——
SPECIAL_TEMPLATE_STEM = "A02办案安全承诺书"              # 不含扩展名
ROLE_PHRASE           = "(办案组组长、办案人员、安全员)"  # 模板里的并列写法
ROLES                 = ["办案组组长", "办案人员", "安全员"]

def iter_paragraphs(doc):
    """遍历 doc 所有段落（正文 + 表格单元格）"""
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

def find_placeholders(doc):
    """返回文档里的全部占位符 set()"""
    ph = set()
    for p in iter_paragraphs(doc):
        text = "".join(r.text for r in p.runs)
        ph.update(PLACEHOLDER_RE.findall(text))
    return ph

def replace_in_paragraph(p, mapping):
    """用 mapping 替换段落中的占位符（保留 run 结构）"""
    old = "".join(r.text for r in p.runs)
    new = old
    for k, v in mapping.items():
        new = new.replace(f"`{k}`", v)
    if new != old:
        runs = p.runs or [p.add_run("")]
        runs[0].text = new
        for r in runs[1:]:
            r.text = ""

def replace_everywhere(doc, mapping):
    for p in iter_paragraphs(doc):
        replace_in_paragraph(p, mapping)

def replace_role_phrase(doc, role):
    """把 ROLE_PHRASE 替换成某个具体身份"""
    for p in iter_paragraphs(doc):
        old = "".join(r.text for r in p.runs)
        new = old.replace(ROLE_PHRASE, role)
        if new != old:
            runs = p.runs or [p.add_run("")]
            runs[0].text = new
            for r in runs[1:]:
                r.text = ""

# =================== 1. 扫描模板，汇总占位符 ===================
all_templates = [p for p in TEMPLATE_DIR.rglob("A*.docx") if p.is_file()]
if not all_templates:
    messagebox.showerror("未找到模板", "未在目录及子目录下发现 A*.docx 模板文件。")
    raise SystemExit

all_ph = set()
for tpl in all_templates:
    all_ph.update(find_placeholders(Document(tpl)))

defaults = {
    "填表日期": datetime.date.today().strftime("%Y-%m-%d"),
}

# =================== 2. 收集占位符的值 ===================
root = tk.Tk(); root.withdraw()
mapping = {}

for key in sorted(all_ph):
    init = defaults.get(key, "")
    val = simpledialog.askstring("填写模板字段", f"请输入「{key}」：", initialvalue=init)
    if val is None:
        messagebox.showinfo("已取消", "操作已取消。")
        raise SystemExit
    mapping[key] = val

# ---- 把“安全首课日期”转为“YYYY年MM月DD日” ----
def normalize_date(text):
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            d = datetime.datetime.strptime(text.strip(), fmt).date()
            return f"{d.year}年{d.month:02d}月{d.day:02d}日"
        except ValueError:
            continue
    return text

if "安全首课日期" in mapping:
    mapping["安全首课日期"] = normalize_date(mapping["安全首课日期"])

# ---- 创建输出文件夹 ----
obj_name   = mapping.get("对象姓名", "未命名")
output_dir = TEMPLATE_DIR / f"{obj_name}-走读式谈话"
output_dir.mkdir(exist_ok=True)

# =================== 3. 生成新文档 ===================
generated = []

for tpl in all_templates:
    # ---------- 特殊模板：A02办案安全承诺书 ----------
    if tpl.stem == SPECIAL_TEMPLATE_STEM:
        for role in ROLES:
            doc = Document(tpl)                    # 每次重新读模板
            replace_everywhere(doc, mapping)       # 占位符替换
            replace_role_phrase(doc, role)         # 身份替换

            # 文件名：占位符替换后 + “-角色”
            fname = tpl.stem
            for k, v in mapping.items():
                fname = fname.replace(f"`{k}`", v)
            fname = f"{fname}-{role}.docx"

            path = output_dir / fname
            doc.save(path)
            generated.append(path)
    # ---------- 其他模板：正常生成 1 份 ----------
    else:
        doc = Document(tpl)
        replace_everywhere(doc, mapping)

        fname = tpl.stem
        for k, v in mapping.items():
            fname = fname.replace(f"`{k}`", v)
        path = output_dir / f"{fname}.docx"

        doc.save(path)
        generated.append(path)

# =================== 4. 完成提示 ===================
msg = "已生成以下文件：\n" + "\n".join(str(p) for p in generated)
messagebox.showinfo("完成", msg)
