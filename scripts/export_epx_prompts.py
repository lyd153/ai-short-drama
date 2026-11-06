# -*- coding: utf-8 -*-
"""从 docs/EP-改版/*.md 提取逐镜四行提示词，导出 prompts/EP-改版-12集-prompts.md / .txt。

用法：python scripts/export_epx_prompts.py
规则：文档中的提示词块形如
    S01｜大远·14mm
    主体动作：……
    镜头：……
    光色：……
    风格：……
集与段的上下文来自文件名（EP0X）与「## D1 ｜ 00:00–00:15 ｜ 标题」标题行。
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT, "docs", "EP-改版")
OUT_MD = os.path.join(ROOT, "prompts", "EP-改版-12集-prompts.md")
OUT_TXT = os.path.join(ROOT, "prompts", "EP-改版-12集-prompts.txt")

FIELDS = ("主体动作：", "镜头：", "光色：", "风格：")
SHOT_RE = re.compile(r"^(S\d+)\s*[｜|]\s*(.+)$")
DUR_RE = re.compile(r"^\|\s*(S\d+)\s*\|\s*(\d+)\s*\|")
UNIT_RE = re.compile(r"^##\s+(D\d+)\s*[｜|]\s*([0-9:–\-]+)\s*[｜|]\s*(.+)$")
EP_RE = re.compile(r"^EP(\d+)-(\d+)秒")
TITLE_RE = re.compile(r"^#\s+(EP\d+)\s*·\s*(.+)$")


def parse_file(path):
    """返回 (ep_no, ep_title, total_seconds, [ (unit_id, unit_time, unit_title, [shot dict]) ])"""
    with open(path, "r", encoding="utf-8") as f:
        lines = [ln.rstrip("\n") for ln in f]
    m = TITLE_RE.match(lines[0]) if lines else None
    ep_no = m.group(1) if m else "EP??"
    ep_title = m.group(2).strip() if m else os.path.basename(path)
    m2 = EP_RE.match(os.path.basename(path))
    total_sec = int(m2.group(2)) if m2 else None

    # 分镜表里的时长（S01 → 4s），用于给提示词块补时长
    dur = {}
    for ln in lines:
        md = DUR_RE.match(ln)
        if md:
            dur[md.group(1)] = int(md.group(2))

    units = []
    cur_unit = None
    i = 0
    while i < len(lines):
        ln = lines[i]
        mu = UNIT_RE.match(ln)
        if mu:
            cur_unit = {"id": mu.group(1), "time": mu.group(2), "title": mu.group(3).strip(), "shots": []}
            units.append(cur_unit)
            i += 1
            continue
        ms = SHOT_RE.match(ln)
        if ms and i + 4 < len(lines) and all(FIELDS[k] in lines[i + 1 + k] for k in range(4)):
            sid, lens = ms.group(1), ms.group(2).strip()
            body = [lines[i + 1 + k].split("：", 1)[1].strip() for k in range(4)]
            if cur_unit is None:
                cur_unit = {"id": "D?", "time": "", "title": "", "shots": []}
                units.append(cur_unit)
            cur_unit["shots"].append({"sid": sid, "lens": lens, "dur": dur.get(sid), "body": body})
            i += 5
            continue
        i += 1
    return ep_no, ep_title, total_sec, units


def clean(s):
    """去掉 Markdown 粗体标记，保证 txt 可直接粘贴。"""
    return s.replace("**", "").strip()


def main():
    files = sorted(
        f for f in os.listdir(SRC_DIR)
        if f.startswith("EP") and f.endswith(".md")
    )
    if not files:
        print("no episode files found", file=sys.stderr)
        return 1

    body = []
    txt = []
    total_shots = 0
    total_sec = 0
    stats = []
    for fn in files:
        ep_no, ep_title, sec, units = parse_file(os.path.join(SRC_DIR, fn))
        ep_title = re.sub(r"（[^（）]*）\s*$", "", ep_title).strip()
        n_shots = sum(len(u["shots"]) for u in units)
        total_shots += n_shots
        total_sec += sec or 0
        stats.append((ep_no, ep_title, sec, n_shots, len(units)))

        body.append(f"\n## {ep_no} · {ep_title}（{sec} 秒 / {len(units)} 段 / {n_shots} 条）\n")
        txt.append(f"===== {ep_no} · {ep_title}（{sec}s / {len(units)} 段 / {n_shots} 条）=====\n")

        for u in units:
            body.append(f"\n### {u['id']} ｜ {u['time']} ｜ {u['title']}\n")
            txt.append(f"--- {ep_no}-{u['id']} ｜ {u['time']} ｜ {u['title']} ---\n")
            for s in u["shots"]:
                label = f"{ep_no}-{u['id']}-{s['sid']}"
                dur = f"{s['dur']}s" if s["dur"] else ""
                body.append(f"\n#### {label} ｜ {dur} ｜ {s['lens']}\n")
                for k, name in enumerate(("主体动作", "镜头", "光色", "风格")):
                    body.append(f"- **{name}**：{clean(s['body'][k])}\n")
                txt.append(f"【{label}｜{dur}｜{s['lens']}】")
                for k, name in enumerate(("主体动作", "镜头", "光色", "风格")):
                    txt.append(f"{name}：{clean(s['body'][k])}")
                txt.append("")

    head = [
        "# 改版 12 集 · 全量提示词汇总\n",
        "> 由 `scripts/export_epx_prompts.py` 从 `docs/EP-改版/*.md` 导出，**不要手改**。\n",
        "> 每条四行：主体动作 / 镜头 / 光色 / 风格；可直接复制到即梦、可灵、Seedance 等工具。\n",
        f"> **合计：{len(stats)} 集 / {total_shots} 条提示词 / {total_sec} 秒"
        f"（{total_sec // 60} 分 {total_sec % 60} 秒）**\n",
    ]

    os.makedirs(os.path.dirname(OUT_MD), exist_ok=True)
    with open(OUT_MD, "w", encoding="utf-8", newline="\n") as f:
        f.write("".join(head + body))
    with open(OUT_TXT, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(txt).replace("\n\n\n", "\n\n"))

    print(f"episodes={len(stats)} shots={total_shots} seconds={total_sec}")
    for ep, t, sec, n, u in stats:
        print(f"  {ep} {t} | {sec}s | {u} 段 | {n} 条")
    print("out:", OUT_MD)
    print("out:", OUT_TXT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
