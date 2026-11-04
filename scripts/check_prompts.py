# -*- coding: utf-8 -*-
"""校验分镜数据：单元时长、总时长、提示词四行、镜号唯一性。

用法（仓库根目录）：
    python scripts/check_prompts.py            # 校验 data/ 下全部 EP*.json
    python scripts/check_prompts.py EP01       # 只校验某一集

退出码：0 = 全部通过；1 = 有失败项（便于接 CI）。
"""
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
UNIT_SEC = 15


def check(path):
    with open(path, encoding="utf-8") as f:
        ep = json.load(f)

    name = ep.get("episode", os.path.basename(path))
    units = ep.get("units", [])
    errors, warns = [], []

    if ep.get("durationSec") != sum(s["durationSec"] for u in units for s in u["shots"]):
        errors.append(f"{name}: durationSec 与镜头时长之和不一致")

    seen = set()
    for u in units:
        total = sum(s["durationSec"] for s in u["shots"])
        if total != UNIT_SEC:
            errors.append(f"{name}/{u['id']}: 单元时长 {total}s ≠ {UNIT_SEC}s")
        if len(u["shots"]) < 3:
            warns.append(f"{name}/{u['id']}: 只有 {len(u['shots'])} 个镜头，可能不足以撑满 15 秒")
        for s in u["shots"]:
            sid = s["id"]
            if sid in seen:
                errors.append(f"{name}: 镜号重复 {sid}")
            seen.add(sid)
            p = s.get("prompt", {})
            for key in ("subject", "camera", "light", "style"):
                if not str(p.get(key, "")).strip():
                    errors.append(f"{name}/{sid}: 提示词缺 {key}")
            if s["durationSec"] <= 0:
                errors.append(f"{name}/{sid}: 时长非法")
        # 钩子必须是一个问句或带疑问/悬念
        if not str(u.get("hook", "")).strip():
            warns.append(f"{name}/{u['id']}: 缺观众钩子")

    return name, len(units), len(seen), errors, warns


def main():
    only = sys.argv[1].upper() if len(sys.argv) > 1 else None
    files = sorted(
        f for f in os.listdir(DATA)
        if f.lower().endswith(".json") and (only is None or f.upper().startswith(only))
    )
    if not files:
        print("没有找到待校验的数据文件")
        return 1

    failed = False
    for f in files:
        name, nu, ns, errors, warns = check(os.path.join(DATA, f))
        status = "PASS" if not errors else "FAIL"
        print(f"[{status}] {name}: {nu} 单元 / {ns} 镜")
        for e in errors:
            print("  ✗", e)
            failed = True
        for w in warns:
            print("  · ", w)
    print("\n结论：", "有失败项" if failed else "全部通过")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
