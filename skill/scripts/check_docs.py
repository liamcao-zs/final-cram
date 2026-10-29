# -*- coding: utf-8 -*-
"""
check_docs.py — 交付前的机械检查（不判断内容对错，只抓结构性错误）

用法:
    python check_docs.py \
      --questions _source/questions.md \
      --answers   _source/answers.md \
      --grandma   _source/notes_grandma.md \
      --muscle    _source/notes_muscle.md

检查项:
    1. 题库与答案版的题号是否一一对应
    2. 答案版每题是否都有答案块
    3. 无答案版是否混入答案痕迹
    4. 选择题的答案字母是否都在题目给出的选项里
    5. 题库与答案版的题干是否逐题一致
    6. 题目之间是否近重复（同考点换数字会被抓出来）
    7. 每章题号是否从 1 开始且连续
    8. 老奶版/薄肌版章节是否对齐
    9. 笔记每章是否有「本章速览」
    10. 引用的图片是否存在，文字说「如图」时是否真配了图
    11. 「耍起」是否超量
    12. 是否残留 BOSS/血条/打怪/必杀技/通关标准
    13. 表格列数是否一致
    14. 源文件头是否缺字段
"""

import argparse
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

FAILS, WARNS, PASSES = [], [], []


def read(p):
    path = Path(p)
    if not path.exists():
        sys.exit(f"找不到文件: {p}")
    return path.read_text(encoding="utf-8")


def h1_titles(text):
    out = []
    for raw in re.findall(r"^#\s+(.+)$", text, re.M):
        title = raw.split("|")[0]
        title = re.sub(r"\{\{.*?\}\}", "", title).strip()
        out.append(title)
    return out


def q_ids(text):
    return re.findall(r"^###\s+(\d+-\d+)\b", text, re.M)


def sections(text):
    """题号 -> 该题到下一题之间的正文"""
    parts = re.split(r"^###\s+", text, flags=re.M)
    out = {}
    for part in parts[1:]:
        m = re.match(r"(\d+-\d+)\b", part)
        if m:
            out[m.group(1)] = part
    return out


def normalize_stem(body):
    """去掉选项、标点、数字，只留汉字，用于近重复检测。"""
    lines = body.split(chr(10))
    if lines and re.match(r"^\d+-\d+", lines[0].strip()):
        lines = lines[1:]          # 去掉题号与标签那一行
    keep = []
    for line in lines:
        if re.match(r"^[A-DＡ-Ｄ][.、．]", line.strip()):
            break
        keep.append(line)
    stem = "".join(keep)
    return "".join(ch for ch in stem if 0x4E00 <= ord(ch) <= 0x9FFF)


def check_duplicates(q_text, threshold=0.86):
    ids = q_ids(q_text)
    secs = sections(q_text)
    stems = {qid: normalize_stem(secs.get(qid, "")) for qid in ids}
    hits = []
    for i in range(len(ids)):
        for j in range(i + 1, len(ids)):
            a, b = stems[ids[i]], stems[ids[j]]
            if len(a) < 12 or len(b) < 12:
                continue
            ratio = SequenceMatcher(None, a, b).ratio()
            if ratio >= threshold:
                hits.append((ids[i], ids[j], ratio))
    warn(not hits,
         "未发现近重复题",
         "疑似重复: " + "; ".join(f"{a}/{b} {r:.2f}" for a, b, r in hits[:8]),
         "题目去重")


def _to_ascii_letter(ch):
    """全角 Ａ-Ｄ 转成 A-D，避免同一选项被当成两种写法。"""
    o = ord(ch)
    return chr(o - 0xFEE0) if 0xFF21 <= o <= 0xFF24 else ch


def _option_letters(text):
    return {_to_ascii_letter(x)
            for x in re.findall(r"^([A-DＡ-Ｄ])[.、．]", text, re.M)}


def check_answer_letters(q_text, a_text):
    """选择题的答案字母必须落在题目实际给出的选项里；非选择题跳过。"""
    qsec, asec = sections(q_text), sections(a_text)
    bad, checked = [], 0
    for qid, body in asec.items():
        opts = _option_letters(qsec.get(qid, ""))
        if not opts:
            continue                    # 计算/简答题，答案本来就不是字母
        nl = chr(10)
        m = re.search(r"::: answer[^" + nl + r"]*" + nl + r"(.+?)" + nl + r":::", body, re.S)
        if not m:
            continue
        checked += 1
        first = next((ln for ln in m.group(1).splitlines() if ln.strip()), "")
        plain = first.replace("*", "").strip()
        head = re.match(r"^([A-Za-zＡ-Ｄ](?:[、,，\s]*[A-Za-zＡ-Ｄ])*)", plain)
        raw = ("".join(re.findall(r"[A-Za-zＡ-Ｄ]", head.group(1))) if head
               else "".join(re.findall(r"[A-DＡ-Ｄ]", plain.split("（")[0])))
        letters = [_to_ascii_letter(x).upper() for x in raw]
        if not letters:
            bad.append(f"{qid}: 未识别到答案字母")
        elif any(x not in opts for x in letters):
            bad.append(f"{qid}: 答案 {''.join(letters)} 超出选项 {''.join(sorted(opts))}")
    check(not bad, f"{checked} 道选择题答案字母都在选项内", "; ".join(bad[:6]), "答案与选项")


def _stem_only(body):
    """题号行之后、选项或答案块之前的部分，去掉所有空白，用于比对题干。"""
    lines = body.splitlines()
    if lines and re.match(r"^\d+-\d+", lines[0].strip()):
        lines = lines[1:]
    keep = []
    for line in lines:
        s = line.strip()
        if (s.startswith("#") or s.startswith("::: answer")
                or re.match(r"^[A-DＡ-Ｄ][.、．]", s)):
            break
        keep.append(line)
    return "".join("".join(keep).split())


def check_stem_match(q_text, a_text):
    """答案版必须原样保留无答案版的题干，防止两边各改各的。"""
    qsec, asec = sections(q_text), sections(a_text)
    bad = []
    for qid in qsec:
        if qid not in asec:
            continue
        q_stem, a_stem = _stem_only(qsec[qid]), _stem_only(asec[qid])
        if q_stem != a_stem:
            bad.append(qid)
    check(not bad, "两版题干逐题一致",
          f"题干被改动: {', '.join(bad[:8])}", "题干一致")


def check_numbering(q_text):
    """每章题号应为 1-1, 1-2 … 连续，中间不跳号。"""
    chapters = re.split(r"^#\s+", q_text, flags=re.M)[1:]
    bad = []
    for chunk in chapters:
        lines = chunk.splitlines()
        title = lines[0].split("|")[0].strip() if lines else ""
        ids = re.findall(r"^###\s+(\d+)-(\d+)\b", chunk, re.M)
        if not ids:
            continue
        nums = [int(b) for _, b in ids]
        if nums != list(range(1, len(nums) + 1)):
            bad.append(f"{title}: {nums}")
    check(not bad, "每章题号从 1 开始且连续",
          "题号不连续: " + "; ".join(bad[:5]), "题号连续")


def check(cond, ok_msg, fail_msg, label):
    if cond:
        PASSES.append(f"{label}: {ok_msg}")
    else:
        FAILS.append(f"{label}: {fail_msg}")


def warn(cond, ok_msg, warn_msg, label):
    if cond:
        PASSES.append(f"{label}: {ok_msg}")
    else:
        WARNS.append(f"{label}: {warn_msg}")


def check_pair(q_path, a_path):
    q_text, a_text = read(q_path), read(a_path)
    qids, aids = q_ids(q_text), q_ids(a_text)
    check(qids == aids,
          f"{len(qids)} 题，题号与顺序一致",
          f"题号不一致。题库 {len(qids)} 题 / 答案 {len(aids)} 题；"
          f"缺失 {sorted(set(qids) - set(aids))}，多余 {sorted(set(aids) - set(qids))}",
          "题号对应")

    missing = [qid for qid, body in sections(a_text).items()
               if qid in aids and "::: answer" not in body]
    check(not missing, "每题都有答案块", f"缺答案块: {missing}", "答案完整性")

    leaks = []
    for i, line in enumerate(q_text.splitlines(), 1):
        s = line.strip()
        if re.match(r":::\s*answer", s):
            leaks.append((i, s))
        elif re.match(r"^(答案|正确答案|参考答案|解析)\s*[:：]", s):
            leaks.append((i, s))
        elif "故选" in s or "正确答案是" in s:
            leaks.append((i, s))
    check(not leaks,
          "没有答案痕迹",
          "疑似答案泄漏 " + "; ".join(f"L{i}: {s[:40]}" for i, s in leaks[:5]),
          "无答案版")


def check_images(paths, base_dir):
    missing = []
    total = 0
    for p in paths:
        p = Path(p)
        if not p.exists():
            continue
        text = read(p)
        for rel in re.findall(r"!\[[^\]]*\]\(([^)]+)\)", text):
            rel = rel.strip().strip('"').strip("'")
            if rel.startswith(("http://", "https://")):
                continue
            total += 1
            target = Path(rel)
            if not target.is_absolute():
                target = Path(base_dir) / target
            if not target.exists():
                missing.append(f"{Path(p).name} -> {rel}")
    check(not missing, f"{total} 张引用图都存在",
          "缺图: " + "; ".join(missing[:5]), "图片")

    # 文字里说「如图」却整篇没有插图，多半是忘了配图
    for p in paths:
        p = Path(p)
        if not p.exists():
            continue
        text = read(p)
        refs = re.findall(r"如图|图示|下图", text)
        if refs and not re.search(r"!\[[^\]]*\]\(", text):
            warn(False, "",
                 f"{p.name}: {len(refs)} 处「如图/图示/下图」但全文无插图",
                 f"配图[{p.name}]")


def check_overview(path):
    """笔记每章开头应有「本章速览」；开篇/总览/提词器这类特殊页除外。"""
    text = read(path)
    chapters = re.split(r"^#\s+", text, flags=re.M)[1:]
    missing = []
    for chunk in chapters:
        lines = chunk.splitlines()
        title = lines[0].split("|")[0].strip() if lines else ""
        if any(k in title for k in ("开篇", "总览", "提词器", "速记")):
            continue
        if "::: overview" not in chunk and "::: boss" not in chunk:
            missing.append(title)
    warn(not missing, f"{Path(path).name}: 每章都有本章速览",
         f"{Path(path).name}: 缺本章速览 {missing[:5]}",
         f"本章速览[{Path(path).name}]")


def check_notes(g_path, m_path):
    g, m = h1_titles(read(g_path)), h1_titles(read(m_path))
    check(g == m, f"两版章节一致（{len(g)} 章）",
          f"章节不一致：仅老奶版 {sorted(set(g) - set(m))}，仅薄肌版 {sorted(set(m) - set(g))}",
          "双版对齐")


def check_playful(paths):
    for p in paths:
        p = Path(p)
        if not p.exists():
            continue
        text = read(p)
        n = text.count("耍起")
        chapters = max(len(h1_titles(text)), 1)
        warn(n <= chapters,
             f"出现 {n} 次 / {chapters} 章",
             f"出现 {n} 次 / {chapters} 章，超过「每章最多一次」",
             f"耍起次数[{p.name}]")


def check_game_words(paths):
    words = ["BOSS", "血条", "打怪", "必杀技", "通关标准"]
    hits = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            continue
        for i, line in enumerate(read(p).splitlines(), 1):
            for w in words:
                if w in line:
                    hits.append(f"{p.name}:L{i} {w}")
    check(not hits, "无游戏化措辞", "残留: " + "; ".join(hits[:8]), "措辞")


def check_tables(paths):
    problems = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            continue
        block, start = [], 0

        def flush():
            nonlocal block, start
            if len(block) >= 2:
                counts = {len(re.findall(r"(?<!\\)\|", line)) for _, line in block}
                if len(counts) > 1:
                    problems.append(f"{p.name}:L{start} 列数不一致 {sorted(counts)}")
            block = []

        for i, line in enumerate(read(p).splitlines(), 1):
            if line.strip().startswith("|"):
                if not block:
                    start = i
                if not re.fullmatch(r"\|[\s:|-]+\|", line.strip()):
                    block.append((i, line))
            else:
                flush()
        flush()
    check(not problems, "表格列数一致", "; ".join(problems[:5]), "表格")


def check_front_matter(paths):
    need = {"title", "course", "track", "date"}
    bad = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            continue
        text = read(p)
        if not text.startswith("---"):
            bad.append(f"{p.name}: 缺 front matter")
            continue
        end = text.find("\n---", 3)
        meta = text[3:end] if end != -1 else ""
        keys = {line.split(":", 1)[0].strip() for line in meta.splitlines() if ":" in line}
        miss = need - keys
        if miss:
            bad.append(f"{p.name}: 缺 {sorted(miss)}")
    check(not bad, "源文件头完整", "; ".join(bad), "front matter")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--questions")
    ap.add_argument("--answers")
    ap.add_argument("--grandma")
    ap.add_argument("--muscle")
    ap.add_argument("--base-dir", default=".", help="解析图片相对路径的基准目录")
    ap.add_argument("--dup-threshold", type=float, default=0.86,
                    help="题目近重复阈值，0~1，默认 0.86")
    args = ap.parse_args()

    if args.questions and args.answers:
        q_text, a_text = read(args.questions), read(args.answers)
        check_pair(args.questions, args.answers)
        check_answer_letters(q_text, a_text)
        check_stem_match(q_text, a_text)
    if args.questions:
        check_duplicates(read(args.questions), args.dup_threshold)
        check_numbering(read(args.questions))

    if args.grandma and args.muscle:
        check_notes(args.grandma, args.muscle)
    for p in dict.fromkeys(p for p in (args.grandma, args.muscle) if p):
        check_overview(p)

    all_paths = [p for p in (args.questions, args.answers, args.grandma, args.muscle) if p]
    check_images(all_paths, args.base_dir)
    playful_paths = list(dict.fromkeys(p for p in (args.questions, args.grandma, args.muscle) if p))
    check_playful(playful_paths)
    check_game_words(all_paths)
    check_tables(all_paths)
    check_front_matter(all_paths)

    for line in PASSES:
        print(f"[ OK ] {line}")
    for line in WARNS:
        print(f"[WARN] {line}")
    for line in FAILS:
        print(f"[FAIL] {line}")

    print(f"\n通过 {len(PASSES)} / 警告 {len(WARNS)} / 失败 {len(FAILS)}")
    sys.exit(1 if FAILS else 0)


if __name__ == "__main__":
    main()
