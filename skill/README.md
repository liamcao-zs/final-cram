# 期末速通 skill 使用说明

## 怎么触发

把课件、作业、真题（PDF / PPT / Word）丢给 Claude，然后说一句：

> 用 final-cram 帮我速通这门课，还有 6 小时。

或者直接用触发词：**期末速通 / 考前突击 / 老奶版 / 薄肌版 / 按章节通关**。

## 它会先问你三件事

1. 考到哪、什么题型
2. 还有几小时
3. 哪几章最虚

不想回答就说"你定"，它按材料自动估。

## 你会拿到什么

| 文件 | 内容 |
| --- | --- |
| `01_速通笔记_老奶版.docx` | 零基础也能看懂，生活比喻为主 |
| `01_速通笔记_薄肌版.docx` | 有基础快速过，结论 + 步骤 + 陷阱 |
| `02_通关题库_无答案.docx` | 按章节排的题，去重后的 |
| `03_通关题库_答案解析.docx` | 同题，每题三段：答案 / 老奶解析 / 薄肌秒杀 |

Word 里：目录可点击跳转，每章末尾可返回目录，四角有你的小人图，页脚有页码。

交付前 `check_docs.py` 会自动检查 14 项：题号对应、答案完整性、答案字母、答案泄漏、题干一致、题号连续、近重复题、缺图、两版对齐、本章速览、耍起次数、游戏化措辞、表格列数等。

## 自己改内容

中间文件在 `_source/`，改完重跑：

```bash
python ~/.claude/skills/final-cram/scripts/build_docx.py \
  --src _source/notes_muscle.md --out 01_速通笔记_薄肌版.docx
```

## 环境要求

Windows + Python 3.10+，依赖：

```bash
pip install -r requirements.txt
```

`verify_docx.py` 需要本机安装 Microsoft Word（用 COM 导 PDF 做验收）；没有 Word 时跳过验收步骤即可。

---

## 主要文件

```
SKILL.md                  流程总纲（Claude 读）
references/
  intake.md               材料摄取与学习边界
  chapter-map.md          章节优先级、时间预算、本章速览
  dual-voice.md           老奶版 / 薄肌版写法
  exam-profiles.md        理科/工科/文科/医学法学四套打法
  kill-methods.md         数学/电路/物理/自控秒杀招式
  question-bank.md        去重选题与双解析
  docx-spec.md            Word 源文件语法与排版
scripts/
  extract_materials.py    PDF/PPTX/DOCX → Markdown + 整页图 + 图片索引
  build_docx.py           生成 Word（目录、标签、本章速览、四角水印）
  check_docs.py           结构自检（题号、答案泄漏、缺图、两版对齐、近重复题）
  verify_docx.py          Word → PDF → PNG 视觉验收
  crop_image.py           从整页截图中裁题图（--grid 出坐标网格）
  make_assets.py          重新生成水印素材
examples/                 示例源文件
assets/mascot.jpg         水印原图
assets/icon.png           512×512 市场图标
```
