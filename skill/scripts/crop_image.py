# -*- coding: utf-8 -*-
"""
crop_image.py — 从整页截图里裁出题图（矢量图/无法直接抽取的图）

用法:
    python crop_image.py --src page-3.png --box 10,20,60,55 --out fig.png
    python crop_image.py --src page-3.png --box 10%,15%,90%,60% --out fig.png
    python crop_image.py --src page-3.png --grid --out page-3-grid.png

box 顺序: 左, 上, 右, 下；纯数字为像素，带 % 为百分比。
"""

import argparse
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
from PIL import Image, ImageDraw, ImageFont


def parse_box(box: str, w: int, h: int):
    vals = [v.strip() for v in box.split(",")]
    if len(vals) != 4:
        raise ValueError("box 需要 4 个值：左,上,右,下")
    out = []
    for v in vals:
        if v.endswith("%"):
            out.append(float(v[:-1]) / 100)
        else:
            out.append(None)
    if all(v is not None for v in out):
        return tuple(int(v * (w if i % 2 == 0 else h)) for i, v in enumerate(out))
    return tuple(int(float(v)) for v in vals)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--box", default=None)
    ap.add_argument("--out", required=True)
    ap.add_argument("--padding", type=int, default=12)
    ap.add_argument("--grid", action="store_true",
                    help="只输出带 10%% 网格和百分比坐标的参考图，不裁切")
    args = ap.parse_args()
    if not args.grid and not args.box:
        ap.error("--box 必填（除非用 --grid）")

    img = Image.open(args.src).convert("RGB")
    if args.grid:
        draw = ImageDraw.Draw(img)
        font = ImageFont.load_default()
        w, h = img.size
        step = max(1, min(w, h) // 500)
        for i in range(1, 10):
            x = int(w * i / 10)
            y = int(h * i / 10)
            draw.line([(x, 0), (x, h)], fill=(220, 40, 40), width=step)
            draw.line([(0, y), (w, y)], fill=(220, 40, 40), width=step)
            draw.text((x + 3, 3), str(i * 10) + "%", fill=(220, 40, 40), font=font)
            draw.text((3, y + 3), str(i * 10) + "%", fill=(220, 40, 40), font=font)
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        img.save(out)
        print("OK grid " + str(out) + "  " + str(w) + "x" + str(h))
        return

    l, t, r, b = parse_box(args.box, img.width, img.height)
    l, t = max(0, l - args.padding), max(0, t - args.padding)
    r, b = min(img.width, r + args.padding), min(img.height, b + args.padding)
    if r <= l or b <= t:
        raise SystemExit(f"裁剪框无效: {(l, t, r, b)}")
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    img.crop((l, t, r, b)).save(out)
    print(f"OK {out}  {r - l}x{b - t}")


if __name__ == "__main__":
    main()
