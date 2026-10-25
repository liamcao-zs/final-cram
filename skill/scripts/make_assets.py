# -*- coding: utf-8 -*-
"""把原图裁白边、缩小、做淡，生成四角水印素材。"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
src = ROOT / "assets" / "mascot.jpg"
dst = ROOT / "assets" / "mascot_faded.png"

img = Image.open(src).convert("RGBA")

# 去掉接近纯白的边，保留主体
gray = img.convert("L")
mask = gray.point(lambda v: 255 if v < 247 else 0)
bbox = mask.getbbox()
if bbox:
    img = img.crop(bbox)

side = 160  # 1.35cm @ 300dpi
img = img.resize((side, side), Image.LANCZOS)

rgb = img.convert("RGB")
out = Image.new("RGBA", img.size)
out.putdata([(r, g, b, 88) for r, g, b in rgb.getdata()])  # 34% 不透明度
out.save(dst, "PNG")
print(f"OK {dst} {out.size}")
