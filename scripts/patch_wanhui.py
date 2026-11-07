# -*- coding: utf-8 -*-
"""给挽回线 md 打补丁：
1) EP01：每镜台词格追加「【特写：…】」；
2) EP02–EP12：每段末镜（S04）追加一句补白台词，提高对话密度。
用法：python scripts/patch_wanhui.py
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "docs", "EP-挽回线")

# ---- EP01 特写（镜号 -> 特写描述）----
CLOSEUPS = {
    "S01": "剑刃横在颈侧，刃面映出她半张脸",
    "S02": "他半睁的眼，睫毛压着烛影",
    "S03": "她握剑的手，指节与剑柄",
    "S04": "他眼底的冷光，落在泪痣上",
    "S01b": "",
}
EP01_CLOSEUPS = [
    ("D1", {"S01": "剑刃横在颈侧，刃面映出她半张脸",
            "S02": "他半睁的眼，睫毛压着烛影",
            "S03": "她握剑的手，指节与剑柄",
            "S04": "他眼底的冷光，落在泪痣上"}),
    ("D2", {"S01": "她俯身时垂下的发丝，扫过他的脸",
            "S02": "他的喉结，剑刃压出的浅痕",
            "S03": "剑刃与颈侧皮肤之间那一线",
            "S04": "他笑时唇角，笑意没到眼睛"}),
    ("D3", {"S01": "他两指捏住剑身的那一下",
            "S02": "刃口与指腹之间一道浅红痕（不出血）",
            "S03": "她撑在他胸口的手，指节泛白",
            "S04": "她红到耳后的侧脸"}),
    ("D4", {"S01": "桌上被顺手推到一边的药碗",
            "S02": "她握剑不放的手",
            "S03": "两人擦身而过时的肩线",
            "S04": "推到她面前的茶碗，水纹荡开"}),
    ("D5", {"S01": "碗沿与两只手接触的地方",
            "S02": "他低头看她的指尖",
            "S03": "她耳根的红，一层层漫上来",
            "S04": "他眼里慢慢漫开的笑意"}),
    ("D6", {"S01": "他敲在桌面上的指节",
            "S02": "她的剑尖在地上点的两下",
            "S03": "他说「还要问几个问题」时的唇形",
            "S04": "剑入鞘的那一瞬，木鞘与刃的摩擦"}),
    ("D7", {"S01": "她停在门边、没有推门的手",
            "S02": "他放下茶杯的动作",
            "S03": "她回头时竖起来又软下去的眼",
            "S04": "他念那个名字时的唇"}),
    ("D8", {"S01": "被带上的门，缝里一线光",
            "S02": "他端起碗又放下的手",
            "S03": "窗棂的影子投在他脸上，像一道锁",
            "S04": "纯黑里浮出的字幕"}),
]

# ---- EP02–EP12：每段末镜补白一句（集 -> 段 -> 台词）----
EXTRA = {
    "EP02": {
        "D1": "林（冷、盯着他的眼）：「你眼里没半点歉意。」",
        "D2": "邶（低声）：「我记得。她撑着的时候，我没让你把草全拿走。」",
        "D3": "林（别开脸）：「你别把命递得这么顺口。」",
        "D4": "邶（对门外、更沉）：「三日内，她要什么，都给。」",
    },
    "EP03": {
        "D1": "林（蹲下时自语）：「这世道，连倒都要倒在我眼前。」",
        "D2": "林（低声）：「我数三下。三下你不睁眼，我就走。」",
        "D3": "邶（昏沉、气声）：「……手……别放。」",
        "D4": "林（把他往背上送）：「你欠我的，记着。别死。」",
    },
    "EP04": {
        "D1": "林（对里屋）：「阿婆，撑住。我把人抬进来就来。」",
        "D2": "邶（看她端碗）：「先喂她。我不急。」",
        "D3": "林（嘴硬）：「你这条命，是拿我阿婆的草换的。」",
        "D4": "林（站在门口喊）：「你要死，别死在我看不见的地方！」",
    },
    "EP05": {
        "D1": "邶（低声）：「三日。我欠你三日。」",
        "D2": "林（昏迷、气声）：「……草……别给兔子。」",
        "D3": "邶（笨手笨脚）：「别动。药汁要流到耳后了。」",
        "D4": "林（半梦半醒）：「你坐着……别走那么远。」",
    },
    "EP06": {
        "D1": "林（看着纸、念）：「一半碾了，一半给阿婆。——他认得药。」",
        "D2": "萧（温和）：「你的脉象跟着山走，不跟着人走。」",
        "D3": "林（低声）：「他要是想取，第一天就取了。」",
        "D4": "萧（笑）：「他站在门外多久，就犹豫了多久。」",
    },
    "EP07": {
        "D1": "林（跑着喊）：「你站住——你欠我的还没还！」",
        "D2": "林（盯着他）：「全部。那你为什么还把草送回来。」",
        "D3": "邶（低声、像自语）：「我这条命，本来就不该活这么久。」",
        "D4": "林（额头抵着他）：「你要真觉得不值，就别闭眼。」",
    },
    "EP08": {
        "D1": "陆驰（低声）：「古地那边，连夜送来的。主子，属下不敢拆。」",
        "D2": "邶（很轻）：「她若失魂，我活着做什么。」",
        "D3": "林（在门口、轻声）：「你在烧东西。我闻见纸灰了。」",
        "D4": "邶（背对她）：「睡吧。明天我去把草籽种上。」",
    },
    "EP09": {
        "D1": "林（坐下时）：「你今晚不对劲。」",
        "D2": "邶（看着月亮）：「我算过所有人的命。唯独没算到有人会给我送草。」",
        "D3": "林（很轻）：「那你算算，我还能陪你几次十五。」",
        "D4": "邶（独白）：「她问我的时候，我一个字都不敢答。」",
    },
    "EP10": {
        "D1": "萧（温和）：「我不说谎。我只挑你最该听的那半句说。」",
        "D2": "林（很稳）：「你不用答。你第一天就知道，这就够了。」",
        "D3": "邶（低声）：「洛筱。我后来……」",
        "D4": "陆驰（急）：「主子，那是您三天的量！」",
    },
    "EP11": {
        "D1": "头目（粗）：「族长说了，活的最好。别弄坏了。」",
        "D2": "邶（很低）：「鞋在这儿。人，往哪去了。」",
        "D3": "族长（笑）：「你拿什么换？你那条命，早就不值钱了。」",
        "D4": "邶（气声）：「她说别走……我答应过她。」",
    },
    "EP12": {
        "D1": "林（抓人问）：「他还有气吗。你说话。」",
        "D2": "林（哭骂）：「你连命都能骗，你还有什么不敢骗的。」",
        "D3": "邶（很轻）：「草还你。命……也还你。」",
        "D4": "林（额头抵着他）：「我不还。你这条命，归我管。」",
    },
}

ROW_RE = re.compile(r"^(\|\s*(S\d+)\s*\|\s*\d+\s*\|[^|]*\|[^|]*\|[^|]*\|)([^|]*)(\|\s*)$")


def patch_ep01(path):
    with open(path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    cur_unit = None
    changed = 0
    out = []
    for ln in lines:
        mu = re.match(r"^##\s*D(\d+)\s*[｜|]", ln)
        if mu:
            cur_unit = f"D{mu.group(1)}"
        mr = ROW_RE.match(ln.rstrip("\n"))
        if mr and cur_unit:
            close = dict(EP01_CLOSEUPS).get(cur_unit, {}).get(mr.group(2), "")
            if close and "【特写" not in mr.group(3):
                cell = mr.group(3).rstrip() + f" 【特写：{close}】 "
                out.append(mr.group(1) + cell + mr.group(4) + "\n")
                changed += 1
                continue
        out.append(ln)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.writelines(out)
    return changed


def patch_other(path, ep):
    table = EXTRA.get(ep)
    if not table:
        return 0
    with open(path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    cur_unit = None
    changed = 0
    out = []
    for ln in lines:
        mu = re.match(r"^##\s*D(\d+)\s*[｜|]", ln)
        if mu:
            cur_unit = f"D{mu.group(1)}"
        mr = ROW_RE.match(ln.rstrip("\n"))
        if mr and cur_unit and mr.group(2) == "S04":
            extra = table.get(cur_unit, "")
            if extra and "【补】" not in mr.group(3):
                cell = mr.group(3).rstrip()
                sep = "／" if cell.strip() not in ("无", "") else ""
                cell = f"{cell}{sep}{extra} 【补】 "
                out.append(mr.group(1) + cell + mr.group(4) + "\n")
                changed += 1
                continue
        out.append(ln)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.writelines(out)
    return changed


def main():
    total = 0
    for fn in sorted(os.listdir(SRC)):
        if not fn.endswith(".md") or not fn.startswith("EP"):
            continue
        path = os.path.join(SRC, fn)
        m = re.match(r"^EP(\d+)-", fn)
        if not m:
            continue
        ep = f"EP{m.group(1).zfill(2)}"
        if ep == "EP01":
            n = patch_ep01(path)
        else:
            n = patch_other(path, ep)
        total += n
        print(f"{ep}: patched {n}")
    print("total:", total)
    return 0


if __name__ == "__main__":
    sys.exit(main())
