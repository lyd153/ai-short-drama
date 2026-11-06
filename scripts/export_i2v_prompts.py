# -*- coding: utf-8 -*-
"""从 docs/EP-改版/*.md 生成「图生视频（image-to-video）」专用提示词。

与四行文生视频提示词的区别：
- 首帧图决定构图与光色，提示词不再写景别/光色/风格；
- 只写：人物（与图一致）、动作、微表情、台词与口型、运镜幅度、禁止项。

关键：分镜表每段（D1…D8）各有一份 S01–S04，必须按「段」分别取行，否则后段覆盖前段。
输出：prompts/EP-改版-12集-图生视频.md / .txt
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT, "docs", "EP-改版")
OUT_MD = os.path.join(ROOT, "prompts", "EP-改版-12集-图生视频.md")
OUT_TXT = os.path.join(ROOT, "prompts", "EP-改版-12集-图生视频.txt")

TITLE_RE = re.compile(r"^#\s+(EP\d+)\s*·\s*(.+)$")
EP_RE = re.compile(r"^EP(\d+)-(\d+)秒")
UNIT_RE = re.compile(r"^##\s+(D\d+)\s*[｜|]\s*([0-9:–\-]+)\s*[｜|]\s*(.+)$")
SHOT_RE = re.compile(r"^(S\d+)\s*[｜|]\s*(.+)$")
ROW_RE = re.compile(r"^\|\s*(S\d+)\s*\|\s*(\d+)\s*\|([^|]*)\|([^|]*)\|([^|]*)\|([^|]*)\|")
FIELDS = ("主体动作：", "镜头：", "光色：", "风格：")

COMMON_NEG = "画面抖动、人物换脸、五官漂移、手指畸形、多余肢体、衣料穿模、背景糊化、速度忽快忽慢"

# 角色识别：全名 / 特征物 / 台词前缀
CAST_SIGNALS = [
    ("邶峙珩", ("邶峙珩", "泪痣", "墨蓝暗纹", "墨蓝锦袍", "墨玉戒")),
    ("林洛筱", ("林洛筱", "银铃", "眉心花钿", "浅蓝对襟", "浅蓝裙")),
    ("陆驰", ("陆驰", "玄色劲装", "眉骨旧疤")),
    ("萧漓", ("萧漓", "白衣男子", "左脸", "百合纹")),
    ("巫罗·岜沅", ("岜沅", "少年")),
    ("青云嫂子", ("青云嫂子",)),
    ("阿婆", ("阿婆",)),
    ("村民", ("村民",)),
    ("小卡（灵兔）", ("灵兔", "兔子", "小卡", "铜铃")),
    ("大祭司", ("大祭司", "老者")),
    ("对立族长/使者", ("族长", "使者")),
]
LINE_SPEAKER = re.compile(r"^(邶|林|陆驰|萧|岜沅|阿婆|村民[甲乙]?|族长|使者|青云嫂子)")

CAST_NEG = {
    "林洛筱": "银铃数量变化、眉心花钿消失、衣色变深",
    "邶峙珩": "泪痣消失或移位、发色变红",
    "萧漓": "左脸百合纹变艳、白衣变灰",
    "小卡（灵兔）": "兔子拟人化、缺角耳复原、铜铃消失",
}


def clean(s):
    return s.replace("**", "").strip()


def detect_cast(*texts):
    text = " ".join(t for t in texts if t)
    found = []
    for name, keys in CAST_SIGNALS:
        if any(k in text for k in keys):
            found.append(name)
    # 台词前缀补充
    for t in texts:
        if not t:
            continue
        m = LINE_SPEAKER.match(t)
        if m:
            mapping = {
                "邶": "邶峙珩", "林": "林洛筱", "陆驰": "陆驰", "萧": "萧漓",
                "岜沅": "巫罗·岜沅", "阿婆": "阿婆", "族长": "对立族长/使者",
                "使者": "对立族长/使者", "青云嫂子": "青云嫂子",
            }
            name = mapping.get(m.group(1))
            if name and name not in found:
                found.append(name)
        if m and m.group(1) in ("村民甲", "村民乙", "村民"):
            if "村民" not in found:
                found.append("村民")
    return found


def parse_ep(path):
    with open(path, "r", encoding="utf-8") as f:
        lines = [ln.rstrip("\n") for ln in f]
    m = TITLE_RE.match(lines[0]) if lines else None
    ep_no = m.group(1) if m else "EP??"
    title = re.sub(r"（[^（）]*）\s*$", "", m.group(2)).strip() if m else ""
    m2 = EP_RE.match(os.path.basename(path))
    total = int(m2.group(2)) if m2 else None

    units = []
    cur = {"id": "D?", "time": "", "title": "", "shots": [], "rows": {}}
    units.append(cur)
    for ln in lines:
        mu = UNIT_RE.match(ln)
        if mu:
            cur = {"id": mu.group(1), "time": mu.group(2), "title": mu.group(3).strip(), "shots": [], "rows": {}}
            units.append(cur)
            continue
        mr = ROW_RE.match(ln)
        if mr:
            cur["rows"][mr.group(1)] = {
                "dur": int(mr.group(2)),
                "lens": clean(mr.group(3)),
                "move": clean(mr.group(4)),
                "face": clean(mr.group(5)),
                "line": clean(mr.group(6)),
            }
            continue
        ms = SHOT_RE.match(ln)
        idx = lines.index(ln) if False else None
    # 第二轮：提取提示词块（需要向后看 4 行）
    i = 0
    cur = units[0]
    while i < len(lines):
        ln = lines[i]
        mu = UNIT_RE.match(ln)
        if mu:
            for u in units:
                if u["id"] == mu.group(1) and u["time"] == mu.group(2) and not u["shots"]:
                    cur = u
                    break
            i += 1
            continue
        ms = SHOT_RE.match(ln)
        if ms and i + 4 < len(lines) and all(FIELDS[k] in lines[i + 1 + k] for k in range(4)):
            sid = ms.group(1)
            body = [clean(lines[i + 1 + k].split("：", 1)[1]) for k in range(4)]
            cur["shots"].append({"sid": sid, "lens": ms.group(2).strip(), "body": body})
            i += 5
            continue
        i += 1
    units = [u for u in units if u["shots"]]
    return ep_no, title, total, units


def build_blocks(ep_no, unit):
    out = []
    for s in unit["shots"]:
        sid = s["sid"]
        row = unit["rows"].get(sid, {})
        dur = row.get("dur")
        lens = row.get("lens") or s["lens"]
        face = row.get("face") or ""
        line = row.get("line") or "无"
        move = row.get("move") or s["body"][1]
        action = s["body"][0]
        key = f"{ep_no}-{unit['id']}-{sid}"
        is_black = ("黑场" in lens) or action.strip() in ("纯黑", "纯黑。")

        L = [f"#### {key} ｜ {dur}s ｜ {lens}", ""]
        L.append(f"- **参考图**：{key} 首帧（构图、光色、人物造型以图为准，提示词只写运动）")
        if is_black:
            sub = re.sub(r"^字幕：", "", line)
            L.append("- **人物**：无（黑场）")
            L.append("- **动作**：纯黑，无画面运动；只做字幕淡入与音频")
            L.append(f"- **字幕**：{sub}")
            L.append(f"- **时长**：{dur} 秒，第 1 帧入画外音，末帧前 0.5 秒淡出")
            L.append(f"- **禁止**：出现任何画面内容、出现人物、字幕抖动")
            out.append("\n".join(L))
            continue

        cast = detect_cast(line, face, action)
        partial = ("手", "指", "腕", "袖", "背影", "肩头", "发丝")
        if cast:
            L.append(f"- **人物**：{'、'.join(cast)}（面部与造型严格等同参考图）")
        elif any(k in (action + face) for k in partial):
            L.append("- **人物**：局部入画（手部/衣袖/背影），不出现正脸，避免生成第二张脸")
        else:
            L.append("- **人物**：无人物（纯空镜）")
        L.append(f"- **动作**：{action}")
        if face and face != "无":
            label = "微表情/神态" if cast else "画面细节"
            L.append(f"- **{label}**：{face}")
        if line and line != "无":
            L.append(f"- **台词与语气**：{line}")
            L.append("- **口型**：逐字对齐，唇形幅度小，句中留半拍停顿，不额外加音效")
        else:
            L.append("- **台词**：无（只有环境声与人物呼吸）")
        L.append(f"- **运镜（首帧已定，只做这一种）**：{move}")
        L.append(f"- **幅度与节奏**：{dur} 秒；动作起于第 0.3 秒，末帧前 0.5 秒收住，匀速不加速")
        neg = COMMON_NEG
        for c in cast:
            if c in CAST_NEG:
                neg += "、" + CAST_NEG[c]
        L.append(f"- **禁止**：{neg}")
        out.append("\n".join(L))
    return out


def main():
    files = sorted(f for f in os.listdir(SRC_DIR) if f.startswith("EP") and f.endswith(".md"))
    if not files:
        print("no files", file=sys.stderr)
        return 1

    md = [
        "# 改版 12 集 · 图生视频（image-to-video）提示词\n",
        "> 由 `scripts/export_i2v_prompts.py` 从 `docs/EP-改版/*.md` 导出，**不要手改**。\n",
        "> 用法：先出该镜首帧图 → 复制本文件对应镜号的整块内容 → 粘进可灵/即梦/Seedance/Vidu 的「图生视频」输入框。\n",
        "> 与四行文生词的区别：**构图与光色由首帧图决定**，这里只写动作、微表情、台词口型、运镜幅度、禁止项。\n",
    ]
    txt = []
    n_eps = n_shots = 0
    for fn in files:
        ep_no, title, total, units = parse_ep(os.path.join(SRC_DIR, fn))
        n_eps += 1
        cnt = sum(len(u["shots"]) for u in units)
        n_shots += cnt
        md.append(f"\n## {ep_no} · {title}（{total} 秒 / {len(units)} 段 / {cnt} 条）\n")
        txt.append(f"===== {ep_no} · {title}（{total}s / {cnt} 条）=====\n")
        for u in units:
            md.append(f"\n### {u['id']} ｜ {u['time']} ｜ {u['title']}\n")
            for block in build_blocks(ep_no, u):
                md.append("\n" + block + "\n")
                txt.append(block.replace("#### ", "【", 1).replace(" ｜ ", "｜", 1) + "\n")

    md.insert(4, f"> **合计：{n_eps} 集 / {n_shots} 条**\n")
    os.makedirs(os.path.dirname(OUT_MD), exist_ok=True)
    with open(OUT_MD, "w", encoding="utf-8", newline="\n") as f:
        f.write("".join(md))
    with open(OUT_TXT, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(txt).replace("\n\n\n", "\n\n"))
    print(f"episodes={n_eps} shots={n_shots}")
    print("out:", OUT_MD)
    print("out:", OUT_TXT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
