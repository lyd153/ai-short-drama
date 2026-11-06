# -*- coding: utf-8 -*-
"""把 docs/EP-改版/*.md 排成 Word（版式参照既有分镜脚本：单元卡 + 分镜表 + 提示词块）。

输出：docs/EP-改版/EP-改版-12集-分镜与提示词.docx
"""
import os
import re
import sys

from docx import Document
from docx.shared import Pt, Mm, Inches, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT, "docs", "EP-改版")
OUT = os.path.join(SRC_DIR, "EP-改版-12集-分镜与提示词_高甜对话版.docx")

C_INK, C_GOLD, C_LINE, C_SOFT, C_WHITE = "1C3B4A", "B08A3E", "C3CFD6", "F4F7F8", "FFFFFF"
C_PROMPT = "F2F6F7"

TITLE_RE = re.compile(r"^#\s+(EP\d+)\s*·\s*(.+)$")
EP_RE = re.compile(r"^EP(\d+)-(\d+)秒")
UNIT_RE = re.compile(r"^##\s+(D\d+)\s*[｜|]\s*([0-9:–\-]+)\s*[｜|]\s*(.+)$")
CARD_RE = re.compile(r"^\*\*单元卡\*\*：(.+)$")
SHOT_RE = re.compile(r"^(S\d+)\s*[｜|]\s*(.+)$")
ROW_RE = re.compile(r"^\|\s*(S\d+)\s*\|\s*(\d+)\s*\|([^|]*)\|([^|]*)\|([^|]*)\|([^|]*)\|")
FIELDS = ("主体动作：", "镜头：", "光色：", "风格：")


def clean(s):
    return s.replace("**", "").strip()


def set_run(run, name="宋体", size=9, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for k in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(k), name)


def style_font(style, name, size, bold=False):
    style.font.name = name
    style.font.size = Pt(size)
    style.font.bold = bold
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for k in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(k), name)


def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def para_shade(p, hex_color):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def para_bar(p, color=C_GOLD, sz=18):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(sz))
    left.set(qn("w:space"), "6")
    left.set(qn("w:color"), color)
    pBdr.append(left)
    pPr.append(pBdr)


def table_borders(t, color=C_LINE, sz=4):
    tblPr = t._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{e}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        b.append(el)
    tblPr.append(b)


def cell_margins(t, top=32, bottom=32, left=56, right=56):
    tblPr = t._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def fixed(t, widths):
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    for i, w in enumerate(widths):
        t.columns[i].width = Inches(w)
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Inches(w)


def no_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def repeat_hdr(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def cell_text(cell, lines, name="宋体", size=8, bold=False, color=None, align=0):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if isinstance(lines, str):
        lines = [lines]
    first = True
    for ln in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.1
        set_run(p.add_run(ln), name=name, size=size, bold=bold, color=color)


def heading(doc, text, level=1, new_page=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.keep_with_next = True
    if new_page:
        pf.page_break_before = True
    if level == 1:
        pf.space_before = Pt(12)
        pf.space_after = Pt(6)
        set_run(p.add_run(text), name="微软雅黑", size=15, bold=True, color=C_INK)
        pPr = p._p.get_or_add_pPr()
        b = OxmlElement("w:pBdr")
        bo = OxmlElement("w:bottom")
        bo.set(qn("w:val"), "single")
        bo.set(qn("w:sz"), "8")
        bo.set(qn("w:space"), "4")
        bo.set(qn("w:color"), C_GOLD)
        b.append(bo)
        pPr.append(b)
    else:
        pf.space_before = Pt(7)
        pf.space_after = Pt(3)
        set_run(p.add_run(text), name="微软雅黑", size=11.5, bold=True, color="2F5A6B")
    return p


def meta_line(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.18
    pf.left_indent = Pt(8)
    pf.keep_with_next = True
    para_shade(p, C_SOFT)
    para_bar(p)
    set_run(p.add_run("单元卡："), name="微软雅黑", size=7.0, bold=True, color="2F5A6B")
    set_run(p.add_run(text), name="宋体", size=7.1, color="22323A")


def prompt_block(doc, sid, dur, lens, body):
    head = doc.add_paragraph()
    pf = head.paragraph_format
    pf.space_before = Pt(3)
    pf.space_after = Pt(0)
    pf.left_indent = Pt(8)
    pf.keep_with_next = True
    para_shade(head, C_PROMPT)
    para_bar(head)
    set_run(head.add_run(f"{sid}　｜　{dur}s　｜　{lens}"), name="微软雅黑", size=7.2, bold=True, color=C_GOLD)
    names = ("主体动作", "镜头", "光色", "风格")
    for i, txt in enumerate(body):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.left_indent = Pt(8)
        pf.line_spacing = 1.05
        pf.keep_with_next = i < len(body) - 1
        para_shade(p, C_PROMPT)
        para_bar(p)
        set_run(p.add_run(f"{names[i]}："), name="微软雅黑", size=6.9, bold=True, color="2F5A6B")
        set_run(p.add_run(txt), name="宋体", size=7.0, color="22323A")


def parse_ep(path):
    with open(path, "r", encoding="utf-8") as f:
        lines = [ln.rstrip("\n") for ln in f]
    m = TITLE_RE.match(lines[0])
    ep_no = m.group(1)
    title = re.sub(r"（[^（）]*）\s*$", "", m.group(2)).strip()
    m2 = EP_RE.match(os.path.basename(path))
    total = int(m2.group(2)) if m2 else 0

    units = []
    cur = None
    i = 0
    while i < len(lines):
        ln = lines[i]
        mu = UNIT_RE.match(ln)
        if mu:
            cur = {"id": mu.group(1), "time": mu.group(2), "title": mu.group(3).strip(),
                   "card": "", "rows": [], "prompts": []}
            units.append(cur)
            i += 1
            continue
        mc = CARD_RE.match(ln)
        if mc and cur is not None:
            cur["card"] = clean(mc.group(1))
            i += 1
            continue
        mr = ROW_RE.match(ln)
        if mr and cur is not None:
            cur["rows"].append({
                "sid": mr.group(1),
                "dur": int(mr.group(2)),
                "lens": clean(mr.group(3)),
                "move": clean(mr.group(4)),
                "face": clean(mr.group(5)),
                "line": clean(mr.group(6)),
            })
            i += 1
            continue
        ms = SHOT_RE.match(ln)
        if ms and i + 4 < len(lines) and all(FIELDS[k] in lines[i + 1 + k] for k in range(4)):
            body = [clean(lines[i + 1 + k].split("：", 1)[1]) for k in range(4)]
            if cur is not None:
                cur["prompts"].append({"sid": ms.group(1), "lens": ms.group(2).strip(), "body": body})
            i += 5
            continue
        i += 1
    return ep_no, title, total, units


def main():
    files = sorted(f for f in os.listdir(SRC_DIR) if f.startswith("EP") and f.endswith(".md"))
    if not files:
        print("no files", file=sys.stderr)
        return 1

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width = Mm(297)
    sec.page_height = Mm(210)
    sec.left_margin = sec.right_margin = Mm(15)
    sec.top_margin = Mm(12)
    sec.bottom_margin = Mm(11)
    style_font(doc.styles["Normal"], "宋体", 10.5)

    hp = sec.header.paragraphs[0]
    hp.alignment = 2
    set_run(hp.add_run("《初遇》改版 · 言情向 12 集 ｜ 分镜与提示词"), name="微软雅黑", size=8, color="8C9AA3")

    fp = sec.footer.paragraphs[0]
    fp.alignment = 1
    set_run(fp.add_run("— "), name="微软雅黑", size=8, color="8C9AA3")
    _r = fp.add_run()
    _f1 = OxmlElement("w:fldChar"); _f1.set(qn("w:fldCharType"), "begin")
    _it = OxmlElement("w:instrText"); _it.set(qn("xml:space"), "preserve"); _it.text = "PAGE"
    _f2 = OxmlElement("w:fldChar"); _f2.set(qn("w:fldCharType"), "end")
    for _e in (_f1, _it, _f2):
        _r._r.append(_e)
    set_run(_r, name="微软雅黑", size=8, color="8C9AA3")
    set_run(fp.add_run(" —"), name="微软雅黑", size=8, color="8C9AA3")

    # 封面
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(26)
    set_run(p.add_run("REWRITE · ROMANCE CUT · 12 EPISODES"), name="微软雅黑", size=8.5, bold=True, color=C_GOLD)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    set_run(p.add_run("初遇 · 改版"), name="微软雅黑", size=32, bold=True, color=C_INK)
    p = doc.add_paragraph()
    set_run(p.add_run("12 集 ｜ EP01 = 120 秒，EP02–EP12 = 60 秒 ｜ 15 秒一段 ｜ 208 镜"), name="微软雅黑", size=12, color="4A6675")
    p = doc.add_paragraph()
    set_run(p.add_run("每段：单元卡（功能/情绪/钩子/声音/转场）→ 分镜表（景别焦段·运镜·画面与表演·台词语气）→ 逐镜四行提示词。"
                      "画面与表演一列含人物动作与说话微表情；台词一列含甜言蜜语与语气标注。"), name="宋体", size=10.5, color="5C7A88")

    n_ep = n_shot = 0
    for idx, fn in enumerate(files):
        ep_no, title, total, units = parse_ep(os.path.join(SRC_DIR, fn))
        n_ep += 1
        cnt = sum(len(u["rows"]) for u in units)
        n_shot += cnt
        heading(doc, f"{ep_no} · {title}（{total} 秒 / {len(units)} 段 / {cnt} 镜）", 1, new_page=(idx > 0))

        for u in units:
            heading(doc, f"{u['id']} ｜ {u['time']} ｜ {u['title']}", 2)
            if u["card"]:
                meta_line(doc, u["card"])

            st = doc.add_table(rows=1, cols=6)
            table_borders(st, sz=3)
            cell_margins(st, 20, 20, 48, 48)
            for j, txt in enumerate(("镜号", "时长", "景别 · 焦段", "运镜", "画面与表演（动作·微表情）", "台词 / 语气")):
                cell_text(st.rows[0].cells[j], txt, name="微软雅黑", size=7.0, bold=True, color=C_WHITE, align=1)
                shade(st.rows[0].cells[j], C_INK)
            repeat_hdr(st.rows[0])
            no_split(st.rows[0])
            for r in u["rows"]:
                row = st.add_row()
                no_split(row)
                cell_text(row.cells[0], r["sid"], name="微软雅黑", size=7.2, bold=True, color=C_INK, align=1)
                cell_text(row.cells[1], f"{r['dur']}s", name="微软雅黑", size=7.2, color="2B3B44", align=1)
                cell_text(row.cells[2], r["lens"], size=6.7, color="2B3B44")
                cell_text(row.cells[3], r["move"], size=6.7, color="3A4A52")
                cell_text(row.cells[4], r["face"], size=6.8, color="1F2E36")
                cell_text(row.cells[5], r["line"], size=6.7, color="6B3A2E")
            fixed(st, [0.6, 0.4, 1.0, 1.45, 4.55, 2.45])

            ph = doc.add_paragraph()
            ph.paragraph_format.space_before = Pt(2)
            ph.paragraph_format.space_after = Pt(0)
            ph.paragraph_format.line_spacing = 1.0
            ph.paragraph_format.keep_with_next = True
            set_run(ph.add_run(f"▍{u['id']} 提示词（{len(u['prompts'])} 条）"), name="微软雅黑", size=7.2, bold=True, color=C_GOLD)
            dur_map = {r["sid"]: r["dur"] for r in u["rows"]}
            for pr in u["prompts"]:
                prompt_block(doc, pr["sid"], dur_map.get(pr["sid"], ""), pr["lens"], pr["body"])

    doc.core_properties.title = "初遇 · 改版 12 集 分镜与提示词"
    doc.core_properties.author = "编剧组"
    doc.save(OUT)
    print("saved:", OUT)
    print("episodes:", n_ep, "shots:", n_shot)
    return 0


if __name__ == "__main__":
    sys.exit(main())
