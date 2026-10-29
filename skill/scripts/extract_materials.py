# -*- coding: utf-8 -*-
"""
extract_materials.py — 把课程资料统一转成可读 Markdown + 可直接使用的图片

支持: .pdf .pptx .docx .txt .md .png .jpg .jpeg .bmp
输出:
    <out>/材料清单.md        总览（文件、规模、文本量、图片、缺口）
    <out>/图索引.md          所有图片及其来源，配图时先看这个
    <out>/text/<名>.md       逐页/逐幻灯片文本，带 [p.N] / [slide N] 标记
    <out>/pages/<名>/...     PDF 整页 PNG（矢量图、扫描页）
    <out>/slides/<名>/...    PPT 整页 PNG（PowerPoint 导出，保真）
    <out>/images/<名>/...    文件内嵌原图
    <out>/cache.json         提取缓存

用法:
    python extract_materials.py 资料目录 --out 提取结果
    python extract_materials.py a.pdf b.pptx --out 提取结果 --page-images low
    python extract_materials.py a.pptx --out 提取结果 --no-slide-images
"""

import argparse
import hashlib
import json
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import pymupdf as fitz
from lxml import etree

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
SUPPORTED = {".pdf", ".pptx", ".docx", ".txt", ".md", ".png", ".jpg", ".jpeg", ".bmp"}


# ------------------------------------------------------------------ 工具

def safe_name(path: Path) -> str:
    name = re.sub(r'[\\/:*?"<>|\s]+', "_", path.stem).strip("_")
    return name[:60] or "material"


def write_text(dst: Path, text: str):
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(text, encoding="utf-8")


def save_pdf_image(doc, raw, dst: Path):
    """PDF 内嵌图常带 SMask（透明蒙版）；直接导出会得到黑底，必须合成到白底。"""
    smask = raw.get("smask", 0)
    if not smask:
        save_binary(raw["image"], dst)
        return
    try:
        import io
        from PIL import Image
        base = Image.open(io.BytesIO(raw["image"])).convert("RGB")
        m = doc.extract_image(smask)
        mask = Image.open(io.BytesIO(m["image"])).convert("L")
        if mask.size != base.size:
            mask = mask.resize(base.size)
        bg = Image.new("RGB", base.size, (255, 255, 255))
        bg.paste(base, (0, 0), mask)
        bg.save(dst, "PNG")
    except Exception:
        dst.write_bytes(raw["image"])


def save_binary(data: bytes, dst: Path):
    """写图片；带透明通道的 PNG 铺白底，避免贴进 Word 变黑块。"""
    dst.write_bytes(data)
    if dst.suffix.lower() != ".png":
        return
    try:
        from PIL import Image
        im = Image.open(dst)
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            im = im.convert("RGBA")
            bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
            bg.alpha_composite(im)
            bg.convert("RGB").save(dst, "PNG")
    except Exception:
        pass


def slide_media(names, loader, no: int) -> dict:
    """返回第 no 张幻灯片引用的媒体文件名列表。"""
    rels = f"ppt/slides/_rels/slide{no}.xml.rels"
    if rels not in names:
        return {}
    root = etree.fromstring(loader(rels))
    out = {}
    for rel in root:
        target = rel.get("Target", "")
        if "media/" in target:
            out[rel.get("Id")] = Path(target.replace("../", "")).name
    return out


# ------------------------------------------------------------------ PDF

def do_pdf(path: Path, out: Path, max_images: int, page_mode="low"):
    doc = fitz.open(path)
    page_count = doc.page_count
    name = safe_name(path)
    img_dir = out / "images" / name
    page_dir = out / "pages" / name
    lines = [f"# {path.name}", "",
             "- 类型: PDF",
             f"- 页数: {page_count}",
             f"- 提取时间: {datetime.now():%Y-%m-%d %H:%M}", ""]

    texts, low_pages, figures = [], [], []
    total_chars, embedded = 0, 0
    seen_images = set()

    # 第一遍：文本 + 内嵌位图
    for i, page in enumerate(doc):
        text = page.get_text("text").strip()
        texts.append(text)
        total_chars += len(text)
        if len(text) < 60:
            low_pages.append(i + 1)

        if embedded < max_images:
            try:
                for k, img in enumerate(page.get_images(full=True), 1):
                    if embedded >= max_images:
                        break
                    raw = doc.extract_image(img[0])
                    data = raw.get("image", b"") if raw else b""
                    if (not raw or len(data) > 3_000_000 or len(data) < 4_000
                            or raw.get("width", 0) < 90 or raw.get("height", 0) < 90):
                        continue
                    digest = hashlib.sha1(data).hexdigest()
                    if digest in seen_images:
                        continue
                    seen_images.add(digest)
                    img_dir.mkdir(parents=True, exist_ok=True)
                    out_img = img_dir / f"p{i + 1:03d}_img{k}.{raw['ext']}"
                    save_pdf_image(doc, raw, out_img)
                    embedded += 1
                    try:
                        bb = page.get_image_bbox(img)
                        near = page.get_text("text", clip=__import__("pymupdf").Rect(
                            bb.x0 - 30, bb.y0 - 80, bb.x1 + 30, bb.y1 + 30)).strip()
                    except Exception:
                        near = ""
                    figures.append({
                        "path": f"images/{name}/{out_img.name}",
                        "source": f"PDF 第 {i + 1} 页内嵌图",
                        "snippet": (near or text).replace(chr(10), " ")[:60],
                    })
            except Exception:
                pass

    # 第二遍：整页导出（矢量图、扫描页只能这样拿）
    if page_mode == "all":
        render_pages = list(range(1, page_count + 1))
    elif page_mode == "low":
        render_pages = low_pages
    else:
        render_pages = []
    render_pages = render_pages[:max_images]

    for no in render_pages:
        page_dir.mkdir(parents=True, exist_ok=True)
        pix = doc[no - 1].get_pixmap(dpi=120)
        f = page_dir / f"p{no:03d}.png"
        pix.save(f)
        figures.append({
            "path": f"pages/{name}/{f.name}",
            "source": f"PDF 第 {no} 页整页",
            "snippet": texts[no - 1].replace(chr(10), " ")[:60],
        })

    rendered = set(render_pages)
    for i, text in enumerate(texts):
        no = i + 1
        marker = f"## [p.{no}]"
        if no in rendered:
            tag = "⚠ 低文本量，" if no in low_pages else ""
            marker += f" {tag}整页图: pages/{name}/p{no:03d}.png"
        lines.append(marker)
        if text:
            lines.append(text)
        lines.append("")

    doc.close()
    coverage = (page_count - len(low_pages)) / max(page_count, 1)
    info = {
        "type": "PDF", "pages": page_count, "chars": total_chars,
        "coverage": f"{coverage:.0%}", "low_pages": low_pages,
        "images": embedded + len(rendered),
        "text_file": f"text/{name}.md", "figures": figures,
    }
    write_text(out / "text" / f"{name}.md", "\n".join(lines))
    return info


# ------------------------------------------------------------------ PPTX

def pptx_to_pdf(path: Path, out: Path):
    """用本机 PowerPoint 把 PPTX 导成 PDF；不可用时返回 None。"""
    cache = out / ".cache"
    cache.mkdir(parents=True, exist_ok=True)
    pdf = cache / f"{safe_name(path)}.pdf"
    if pdf.exists() and pdf.stat().st_mtime >= path.stat().st_mtime:
        return pdf
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return None

    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            app.DisplayAlerts = 1  # ppAlertsNone
        except Exception:
            pass
        pres = app.Presentations.Open(str(path.resolve()), WithWindow=False, ReadOnly=True)
        pres.SaveAs(str(pdf.resolve()), 32)  # ppSaveAsPDF
        pres.Close()
        return pdf if pdf.exists() else None
    except Exception:
        return None
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def do_pptx(path: Path, out: Path, max_images: int, render_slides=True):
    name = safe_name(path)
    text_lines = [f"# {path.name}", "", "- 类型: PPTX", ""]
    total_chars, low_slides, media_count = 0, [], 0
    img_dir = out / "images" / name
    figures = []
    slide_numbers = []
    slide_texts = {}

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        slides = sorted(
            [n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)],
            key=lambda n: int(re.search(r"slide(\d+)\.xml", n).group(1)),
        )
        note_texts = {}
        for n in names:
            m = re.fullmatch(r"ppt/notesSlides/notesSlide(\d+)\.xml", n)
            if m:
                root = etree.fromstring(z.read(n))
                note_texts[int(m.group(1))] = "\n".join(
                    t.text for t in root.iter(f"{{{NS_A}}}t") if t.text and t.text.strip())

        def slide_notes(no: int) -> str:
            """通过 slide 的关系文件找备注；notesSlideN 的编号不保证与页码相同。"""
            rels = f"ppt/slides/_rels/slide{no}.xml.rels"
            if rels not in names:
                return ""
            try:
                rel_root = etree.fromstring(z.read(rels))
            except Exception:
                return ""
            for rel in rel_root:
                if rel.get("Type", "").endswith("/notesSlide"):
                    m = re.search(r"notesSlide(\d+)\.xml", rel.get("Target", ""))
                    if m:
                        return note_texts.get(int(m.group(1)), "")
            return ""

        for s in slides:
            no = int(re.search(r"slide(\d+)\.xml", s).group(1))
            slide_numbers.append(no)
            root = etree.fromstring(z.read(s))
            body = "\n".join(
                t.text for t in root.iter(f"{{{NS_A}}}t") if t.text and t.text.strip()).strip()
            total_chars += len(body)
            slide_texts[no] = body
            if len(body) < 20:
                low_slides.append(no)

            mapping = slide_media(names, z.read, no)
            if mapping and media_count < max_images:
                img_dir.mkdir(parents=True, exist_ok=True)
                for fname in mapping.values():
                    src = f"ppt/media/{fname}"
                    if src not in names:
                        continue
                    dst = img_dir / f"s{no:02d}_{fname}"
                    if not dst.exists():
                        with z.open(src) as fh:
                            save_binary(fh.read(), dst)
                        media_count += 1
                    figures.append({
                        "path": f"images/{name}/{dst.name}",
                        "source": f"PPT 第 {no} 页内嵌图",
                        "snippet": body.replace(chr(10), " ")[:60],
                    })

            text_lines.append(f"## [slide {no}]")
            if body:
                text_lines.append(body)
            note = slide_notes(no)
            if note:
                text_lines.append(f"> 备注: {note.strip()}")
            text_lines.append("")

    # 整页渲染：PPT → PDF → PNG，矢量图/公式/排版全保真
    slide_count = 0
    if render_slides:
        pdf = pptx_to_pdf(path, out)
        if pdf:
            doc = fitz.open(pdf)
            slide_dir = out / "slides" / name
            slide_dir.mkdir(parents=True, exist_ok=True)
            for i in range(min(doc.page_count, max_images)):
                pix = doc[i].get_pixmap(dpi=130)
                f = slide_dir / f"slide-{i + 1:02d}.png"
                pix.save(f)
                real_no = slide_numbers[i] if i < len(slide_numbers) else i + 1
                figures.append({
                    "path": f"slides/{name}/{f.name}",
                    "source": f"PPT 第 {i + 1} 页整页",
                    "snippet": slide_texts.get(real_no, "").replace(chr(10), " ")[:60],
                })
                slide_count += 1
            doc.close()

    info = {
        "type": "PPTX", "slides": len(slides), "chars": total_chars,
        "low_slides": low_slides,
        "images": media_count + slide_count,
        "rendered_slides": slide_count,
        "text_file": f"text/{name}.md", "figures": figures,
    }
    write_text(out / "text" / f"{name}.md", "\n".join(text_lines))
    return info


# ------------------------------------------------------------------ DOCX / 文本 / 图片

def do_docx(path: Path, out: Path, max_images: int):
    from docx import Document
    doc = Document(str(path))
    lines = [f"# {path.name}", "", "- 类型: DOCX", ""]
    chars = 0
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        chars += len(t)
        style = p.style.name if p.style is not None else ""
        prefix = "### " if style.lower().startswith("heading") else ""
        lines.append(prefix + t)
    for ti, table in enumerate(doc.tables, 1):
        lines.append(f"\n## [表格 {ti}]")
        for row in table.rows:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            lines.append("| " + " | ".join(cells) + " |")
            chars += sum(len(c) for c in cells)

    name = safe_name(path)
    img_dir = out / "images" / name
    figures = []
    with zipfile.ZipFile(path) as z:
        media = [n for n in z.namelist() if n.startswith("word/media/")]
        if media:
            img_dir.mkdir(parents=True, exist_ok=True)
        for m in media[:max_images]:
            dst = img_dir / Path(m).name
            with z.open(m) as fh:
                save_binary(fh.read(), dst)
            figures.append({"path": f"images/{name}/{dst.name}", "source": "DOCX 内嵌图"})
    if figures:
        lines.append(f"\n## 内嵌图片 {len(figures)} 张: images/{name}/")

    info = {"type": "DOCX", "chars": chars, "images": len(figures),
            "text_file": f"text/{name}.md", "figures": figures}
    write_text(out / "text" / f"{name}.md", "\n".join(lines))
    return info


def do_text(path: Path, out: Path):
    raw = path.read_text(encoding="utf-8", errors="replace")
    write_text(out / "text" / f"{safe_name(path)}.md", f"# {path.name}\n\n{raw}")
    return {"type": "TEXT", "chars": len(raw),
            "text_file": f"text/{safe_name(path)}.md", "figures": []}


def do_image(path: Path, out: Path):
    name = safe_name(path)
    img_dir = out / "images" / name
    img_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(path, img_dir / path.name)
    rel = f"images/{name}/{path.name}"
    return {"type": "IMAGE", "chars": 0, "images": 1,
            "figures": [{"path": rel, "source": "用户直接提供的图片"}]}


# ------------------------------------------------------------------ 主流程

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("--out", default="提取结果")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--max-images", type=int, default=60)
    ap.add_argument("--page-images", choices=["all", "low", "none"], default="all",
                    help="PDF 整页导出策略：all=全部（默认，配图需要），low=只导低文本页，none=不导")
    ap.add_argument("--no-slide-images", action="store_true",
                    help="不把 PPT 每页渲染成 PNG（默认渲染）")
    args = ap.parse_args()

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    cache_path = out / "cache.json"
    cache = {}
    if cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    files = []
    for item in args.inputs:
        p = Path(item)
        if p.is_dir():
            files += [f for f in sorted(p.rglob("*"))
                      if f.is_file() and f.suffix.lower() in SUPPORTED]
        elif p.is_file():
            files.append(p)

    if not files:
        print("没有找到可处理的文件。支持:", ", ".join(sorted(SUPPORTED)))
        return

    records, skipped = [], 0
    for f in files:
        key = str(f.resolve())
        sig = f"{f.stat().st_mtime_ns}:{f.stat().st_size}"
        if not args.force and cache.get(key, {}).get("sig") == sig:
            records.append(cache[key]["info"] | {"file": f.name})
            skipped += 1
            print(f"[缓存] {f.name}")
            continue

        ext = f.suffix.lower()
        try:
            if ext == ".pdf":
                info = do_pdf(f, out, args.max_images, args.page_images)
            elif ext == ".pptx":
                info = do_pptx(f, out, args.max_images, not args.no_slide_images)
            elif ext == ".docx":
                info = do_docx(f, out, args.max_images)
            elif ext in {".txt", ".md"}:
                info = do_text(f, out)
            else:
                info = do_image(f, out)
            info["file"] = f.name
            records.append(info)
            cache[key] = {"sig": sig, "info": info}
            print(f"[OK] {f.name}  {info}")
        except Exception as e:
            records.append({"file": f.name, "type": ext, "error": str(e), "figures": []})
            print(f"[失败] {f.name}: {e}")

    cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")

    # 材料清单
    md = ["# 材料清单", "",
          f"生成时间: {datetime.now():%Y-%m-%d %H:%M}", "",
          "| 文件 | 类型 | 规模 | 文本量 | 图片 | 备注 |",
          "| --- | --- | --- | --- | --- | --- |"]
    for r in records:
        if "error" in r:
            md.append(f"| {r['file']} | {r['type']} | - | - | - | ❌ {r['error']} |")
            continue
        scale = ""
        if "pages" in r:
            scale = f"{r['pages']} 页"
        elif "slides" in r:
            scale = f"{r['slides']} 张幻灯片"
        note = ""
        if r.get("low_pages"):
            note = f"低文本页 {len(r['low_pages'])} 页，需看图"
        if r.get("low_slides"):
            note = f"低文本幻灯片 {len(r['low_slides'])} 页，需看图"
        if r.get("rendered_slides"):
            note = (note + "；" if note else "") + "已导出整页图"
        md.append(f"| {r['file']} | {r.get('type','')} | {scale} | "
                  f"{r.get('chars', 0)} 字 | {r.get('images', 0)} | {note} |")
    md += ["", "## 下一步", "",
           "1. 先读 `材料清单.md`，判断哪份是主教材、哪份是作业、哪份是真题。",
           "2. 正文不够、但图片数量多的文件，用 Read 工具直接看图。",
           "3. 低文本页/幻灯片必须看图，不要只读文本就下结论。",
           "4. 配图所需素材见 `图索引.md`。"]
    write_text(out / "材料清单.md", "\n".join(md))

    # 图索引
    seen, figs = set(), []
    for r in records:
        for fig in r.get("figures", []):
            if fig["path"] in seen:
                continue
            seen.add(fig["path"])
            figs.append(fig)
    idx = ["# 图索引", "", f"共 {len(figs)} 张。", "",
           "| 图片 | 来源 | 页面文字（前 60 字） |", "| --- | --- | --- |"]
    for fig in figs:
        snippet = (fig.get("snippet") or "").replace("|", "/").replace(chr(10), " ")[:60]
        idx.append(f"| `{fig['path']}` | {fig['source']} | {snippet} |")
    idx += ["", "## 配图流程", "",
            "1. 先看整页图（`slides/` 或 `pages/`），确认图里画了什么。",
            "2. 题目要配图时，优先用原图；带答案标注、红笔、结论箭头的必须裁掉。",
            "3. 整页图裁切：`python crop_image.py --src <整页图> --box 左,上,右,下 --out <图>.png`。",
            "4. 同一张图被多道题共用时，只在第一次出现处插入。",
            "5. 图里的文字和数字必须和题干一致，改过数据的题要重新标注。"]
    write_text(out / "图索引.md", "\n".join(idx))

    print(f"\n完成: {len(records)} 个文件（缓存命中 {skipped}），图片 {len(figs)} 张")
    print(f"清单: {out / '材料清单.md'}")
    print(f"图片: {out / '图索引.md'}")


if __name__ == "__main__":
    main()
