# -*- coding: utf-8 -*-
"""用相位相关估计相邻帧的平移，用多尺度搜索估计推拉，输出每段运镜判断。

用法：python analyze_motion.py <frames_dir> <fps>
帧名要求 t%03d.jpg 连续。
"""
import os
import sys

import numpy as np
from PIL import Image


def load_gray(path, size=(320, 180), y_from=0.0, y_to=1.0):
    im = Image.open(path).convert("L").resize(size, Image.BILINEAR)
    a = np.asarray(im, dtype=np.float64)
    h = a.shape[0]
    a = a[int(h * y_from):int(h * y_to)]
    a -= a.mean()
    w = np.outer(np.hanning(a.shape[0]), np.hanning(a.shape[1]))
    return a * w


def phase_shift(a, b):
    """返回 b 相对 a 的位移 (dy, dx)，单位像素。"""
    fa = np.fft.fft2(a)
    fb = np.fft.fft2(b)
    r = fa * np.conj(fb)
    r /= np.abs(r) + 1e-9
    c = np.fft.ifft2(r).real
    idx = np.unravel_index(np.argmax(c), c.shape)
    dy = idx[0] if idx[0] <= a.shape[0] // 2 else idx[0] - a.shape[0]
    dx = idx[1] if idx[1] <= a.shape[1] // 2 else idx[1] - a.shape[1]
    return dy, dx, c.max()


def zoom_ratio(a, b):
    """粗估缩放：b 若比 a 更「放大」，返回 >1。用中心裁切的相关性搜索。"""
    best, best_s = -1, 1.0
    h, w = a.shape
    for s in (0.94, 0.96, 0.98, 1.0, 1.02, 1.04, 1.06):
        ch, cw = int(h * s), int(w * s)
        y0, x0 = (h - ch) // 2, (w - cw) // 2
        bb = b[y0:y0 + ch, x0:x0 + cw]
        if bb.shape[0] < 8 or bb.shape[1] < 8:
            continue
        bb = np.asarray(Image.fromarray(((bb - bb.min()) / (np.ptp(bb) + 1e-9) * 255).astype(np.uint8)).resize((w, h)), dtype=np.float64)
        bb -= bb.mean()
        bb *= np.outer(np.hanning(h), np.hanning(w))
        _, _, score = phase_shift(a, bb)
        if score > best:
            best, best_s = score, s
    return best_s, best


def main():
    d = sys.argv[1]
    fps = float(sys.argv[2]) if len(sys.argv) > 2 else 5.0
    files = sorted(f for f in os.listdir(d) if f.endswith(".jpg"))
    # 只取画面上方 12%（屋顶/灯笼），尽量排除舞者自身运动
    imgs = [load_gray(os.path.join(d, f), y_from=0.0, y_to=0.12) for f in files]
    print(f"frames={len(files)} fps={fps} duration≈{len(files)/fps:.1f}s")
    print("t(s)   dx     dy     move             zoom")
    prev = imgs[0]
    cum_dx = cum_dy = 0.0
    for i in range(1, len(imgs)):
        dy, dx, score = phase_shift(prev, imgs[i])
        z, zs = zoom_ratio(prev, imgs[i]) if i % 5 == 0 else (1.0, 0)
        cum_dx += dx
        cum_dy += dy
        tag = "static"
        mag = (dx * dx + dy * dy) ** 0.5
        if mag > 3:
            if abs(dx) > abs(dy) * 1.5:
                tag = "pan-L" if dx > 0 else "pan-R"
            elif abs(dy) > abs(dx) * 1.5:
                tag = "tilt"
            else:
                tag = "diag"
        print(f"{i/fps:5.2f}  {dx:5.1f}  {dy:5.1f}  {tag:12s}  {z:.2f}")
        prev = imgs[i]
    print(f"cumulative dx={cum_dx:.1f}px dy={cum_dy:.1f}px (320x180 坐标系)")


if __name__ == "__main__":
    main()
