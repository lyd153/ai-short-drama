# -*- coding: utf-8 -*-
"""把 v2_data.py 中的分镜与提示词导出为 JSON / Markdown / 纯文本三种格式。

用法（仓库根目录）：
    python scripts/export_prompts.py

产物：
    data/EP01.json               机器可读的完整结构（单元 / 镜头 / 提示词）
    prompts/EP01-prompts.md      人读版：按 15 秒单元分组
    prompts/EP01-prompts.txt     粘贴版：每镜一段，可直接丢进生成工具批处理

设计原则：唯一数据源是 scripts/v2_data.py；改动只改它，再跑本脚本重新导出。
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

from v2_data import UNITS, TOTAL_SEC, SHOT_N  # noqa: E402

EPISODE = "EP01"
TITLE = "第一章 · 初遇"


def build_json():
    return {
        "episode": EPISODE,
        "title": TITLE,
        "durationSec": TOTAL_SEC,
        "unitCount": len(UNITS),
        "shotCount": SHOT_N,
        "aspect": "9:16",
        "resolution": "1080x1920",
        "units": [
            {
                "id": u["no"],
                "timecode": u["time"],
                "title": u["title"],
                "function": u["func"],
                "emotion": u["emo"],
                "hook": u["hook"],
                "sound": u["sound"],
                "transition": u["trans"],
                "shots": [
                    {
                        "id": s["sid"],
                        "durationSec": s["dur"],
                        "shotSize": s["lens"],
                        "camera": s["move"],
                        "visual": s["vis"],
                        "dialogue": s["line"],
                        "prompt": {
                            "subject": s["subject"],
                            "camera": s["camera"],
                            "light": s["light"],
                            "style": s["style"],
                        },
                    }
                    for s in u["shots"]
                ],
            }
            for u in UNITS
        ],
    }


def build_md():
    out = []
    out.append(f"# {EPISODE} 提示词库 · {TITLE}\n")
    out.append(f"- 时长：{TOTAL_SEC} 秒（{TOTAL_SEC // 60}′{TOTAL_SEC % 60:02d}″）")
    out.append(f"- 结构：{len(UNITS)} 个 15 秒单元 / {SHOT_N} 个镜头")
    out.append("- 画幅：9:16 1080×1920\n")
    out.append("> 四行固定格式：主体动作 / 镜头 / 光色 / 风格。角色锚与负面词见 "
               "[方法与规范](../docs/00-方法/竖屏短剧分镜与提示词规范.md)。\n")
    for u in UNITS:
        out.append(f"\n## {u['no']} · {u['time']} · {u['title']}\n")
        out.append(f"**功能**：{u['func']}")
        out.append(f"**情绪**：{u['emo']}　|　**钩子**：{u['hook']}\n")
        for s in u["shots"]:
            out.append(f"### {s['sid']} · {s['dur']}s · {s['lens']}")
            out.append("")
            out.append(f"| 项 | 内容 |")
            out.append(f"| --- | --- |")
            out.append(f"| 台词 / 声音 | {s['line']} |")
            out.append(f"| 主体动作 | {s['subject']} |")
            out.append(f"| 镜头 | {s['camera']} |")
            out.append(f"| 光色 | {s['light']} |")
            out.append(f"| 风格 | {s['style']} |")
            out.append("")
    return "\n".join(out)


def build_txt():
    """每镜一段，无 Markdown 装饰，方便批量投喂生成工具。"""
    blocks = []
    for u in UNITS:
        blocks.append(f"# {u['no']} {u['time']} {u['title']}")
        for s in u["shots"]:
            blocks.append(
                f"[{s['sid']}] {s['dur']}s | {s['lens']} | {s['move']}\n"
                f"主体动作：{s['subject']}\n"
                f"镜头：{s['camera']}\n"
                f"光色：{s['light']}\n"
                f"风格：{s['style']}"
            )
    return "\n\n".join(blocks) + "\n"


def write(path, text):
    full = os.path.join(ROOT, path)
    os.makedirs(os.path.dirname(full), exist_ok=True)
    with open(full, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    print(f"wrote {path} ({len(text)} chars)")


if __name__ == "__main__":
    write(f"data/{EPISODE}.json", json.dumps(build_json(), ensure_ascii=False, indent=2))
    write(f"prompts/{EPISODE}-prompts.md", build_md())
    write(f"prompts/{EPISODE}-prompts.txt", build_txt())
