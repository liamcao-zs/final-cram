# -*- coding: utf-8 -*-
"""
verify_docx.py — 用本机 Word 把 .docx 导成 PDF，再逐页转 PNG，做视觉验收。

用法:
    python verify_docx.py --src a.docx --out-dir preview [--pages 1-4] [--dpi 110]

输出:
    preview/a.pdf
    preview/page-1.png ...
    终端打印页数、每页字符数（用于发现空白页 / 溢出页）
"""

import argparse
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import pymupdf as fitz
import pythoncom
import win32com.client


def docx_to_pdf(src: Path, dst: Path):
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(src.resolve()), ReadOnly=True,
                                  AddToRecentFiles=False)
        doc.ExportAsFixedFormat(str(dst.resolve()), 17)  # 17 = wdExportFormatPDF
        doc.Close(False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--pages", default="all")
    ap.add_argument("--dpi", type=int, default=110)
    args = ap.parse_args()

    src = Path(args.src)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf = out_dir / (src.stem + ".pdf")

    docx_to_pdf(src, pdf)
    doc = fitz.open(pdf)
    print(f"PDF: {pdf}  页数: {doc.page_count}")

    if args.pages == "all":
        wanted = range(doc.page_count)
    else:
        a, _, b = args.pages.partition("-")
        start = int(a) - 1
        end = int(b) if b else int(a)
        wanted = range(max(0, start), min(doc.page_count, end))

    for i in wanted:
        page = doc[i]
        text = page.get_text().strip()
        pix = page.get_pixmap(dpi=args.dpi)
        img = out_dir / f"page-{i + 1}.png"
        pix.save(img)
        has_img = bool(page.get_images(full=True))
        flag = ""
        if len(text) < 12:
            flag = "  <-- 空白页?"
        elif len(text) < 50 and not has_img:
            flag = "  <-- 近空白页（只有链接/页码）?"
        print(f"  p{i + 1}: {len(text):4d} chars  {img.name}{flag}")

    doc.close()


if __name__ == "__main__":
    main()
