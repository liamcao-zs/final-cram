---
name: final-cram
version: "1.1.0"
display_name: 期末速通
display_name_en: Final Cram
description_zh: 期末速通资料生成器：读取 PDF、PPT、作业和真题，先划定学习边界，再按章节逐个突破，输出双轨 Word 笔记（老奶版/薄肌版）、无答案通关题库和双解析答案版。
description_en: "A final-exam cram generator: reads PDFs, slides, homework and past papers, sets the learning boundary first, then works chapter by chapter, producing dual-track Word notes (plain-language and quick-review), an answer-free question bank, and a fully explained answer key."
description: 期末速通资料生成器。读取 PDF/PPT/作业/真题，先划定学习边界，再按章节逐个突破，输出老奶版/薄肌版双轨 Word 笔记、无答案通关题库、双解析答案版。Use when 用户提到期末速通、考前突击、几小时复习、章节通关、老奶版/薄肌版、把课件和作业做成复习文档、划重点出题。
---

# 期末速通

把一学期资料压成三份能直接用的 Word：**看得懂的笔记 + 做得完的题 + 抄得走的解析**。

核心假设：用户只剩几个小时，没听过课，目标是拿分，不是学懂。

---

## 交付物

```
01_速通笔记_老奶版.docx
01_速通笔记_薄肌版.docx
02_通关题库_无答案.docx
03_通关题库_答案解析.docx
```

外加 `00_边界卡.md` 和 `_source/`（可重新渲染的源文件）。

---

## 铁律

1. **先定边界，再写内容。** 没输出边界卡之前，不要写正文。
2. **一章一段推进单元。** 每章开头一张「本章速览」：分值、核心、最省时解法、学完能做到、建议用时。**正文里不出现 BOSS、血条、打怪这类字眼。**
3. **老奶版和薄肌版是两份文件。** 同一骨架，不同密度，不混排。
4. **题目和答案物理隔离。** 无答案版里不能出现任何答案痕迹。
5. **每题必须自己算过。** 算不出来就换题，不要交一道自己都没把握的题。
6. **材料没有的内容标 `（AI补充）`。** 不编考点、不编老师原话。
7. **公式用 Unicode，不用 LaTeX。** 见 `references/docx-spec.md`。

---

## 工作流

### Phase 0 · 摄取（5 分钟）

```bash
B=~/.claude/skills/final-cram/scripts   # 本 skill 的 scripts/ 目录
python "$B/extract_materials.py" \
  "资料文件夹" --out "期末速通_课程/_extract"
```

- 先读 `_extract/材料清单.md`。
- 再只读文本的标题层，不要整份灌进上下文。
- 标了"低文本页/幻灯片"的，用 Read 工具看图。
- 配图素材：PPT 每页已导成 PNG（`_extract/slides/`），PDF 内嵌图在 `_extract/images/`，PDF 每页也已整页导出（`_extract/pages/`）。先看 `_extract/图索引.md`，它带“页面文字”，可按题干关键词找图。

详见 [references/intake.md](references/intake.md)。

### Phase 1 · 边界卡（必须先做）

从材料里抽出：章节列表、每章核心方法、前置依赖、可砍内容。然后输出边界卡：

```markdown
# 学习边界卡
## 这次要拿多少分 / 几小时
## 纳入范围（章节 / 知识点数 / 题量）
## 排除范围（明确不看什么，为什么）
## 依据（考纲? 真题? 作业? PPT? 缺口?）
## 输出确认（两版笔记 + 两版题库）
```

同时按 [references/exam-profiles.md](references/exam-profiles.md) 选一套学科模板（理科计算 / 工科专业 / 文科记忆 / 医学法学案例），并写进边界卡。

然后**只问一轮**（每问都有默认值）：

1. 考到哪、什么题型？
2. 还有几小时？
3. 哪几章最虚，老师划重点了吗？

用户说"你定"就按默认走。详见 [references/intake.md](references/intake.md)。

### Phase 2 · 章节地图

按分值 50% + 频率 30% + 可出题性 20% 给每章排优先级，分配时间。

每章开头写一张「本章速览」：分值、核心、最省时解法、学完能做到、建议用时。**这是给读者看的，用正常文档语言，不写游戏化字眼。**

同时把知识点标上 `{{必考}}` / `{{高频}}` / `{{中频}}` / `{{了解}}` / `{{不考}}`。

详见 [references/chapter-map.md](references/chapter-map.md)。

### Phase 3 · 写双轨笔记

按章写 `_source/notes_grandma.md` 和 `_source/notes_muscle.md`。

- 老奶版：生活画面 → 人话定义 → 严格定义 → 考场怎么用。一句一个意思。
- 薄肌版：结论先行 → 条件 → 步骤 → 陷阱 → 考场信号。能表格不段落。

写之前读 [references/dual-voice.md](references/dual-voice.md)（骨架、正反例、禁词、「耍起」规范）和 [references/exam-profiles.md](references/exam-profiles.md)（本学科用哪套字段和答题结构）。

数学/电路/物理题要写秒杀招式时，查 [references/kill-methods.md](references/kill-methods.md)。**每招必须带适用条件**；网络方法先用一道课后题验一遍。

### Phase 4 · 选题 + 双解析

- 先画覆盖矩阵（章节 × 题型），再选题。
- 按**考察点 ID** 去重：同公式换数字算重复。
- 难度配比 基础 40% / 进阶 40% / 拔高 20%。
- 每章开头一句主线，可带一次「耍起」。
- 无答案版只放题；答案版题目完全相同，每题后接三段：答案 / 老奶解析 / 薄肌秒杀。
- 题目有图的必须配图：`![图 1-1 电路图](fig/q1-3.png)`；同一张图多题共用时只插一次。
- 优先顺序：PPT 整页图里裁 > PDF 内嵌图 > 教材截图 > 自己画。带答案标注的图必须裁掉。

详见 [references/question-bank.md](references/question-bank.md)。

### Phase 5 · 自检 + 渲染 + 验收

```bash
# $B 指本 skill 的 scripts/ 目录；装在别处就把这里改成实际路径
B=~/.claude/skills/final-cram/scripts

# 先跑结构自检：题号对应、题干一致、答案泄漏、缺图、两版对齐、游戏化措辞
python "$B/check_docs.py" \
  --questions _source/questions.md --answers _source/answers.md \
  --grandma _source/notes_grandma.md --muscle _source/notes_muscle.md \
  --base-dir _source

python "$B/build_docx.py" --src _source/notes_grandma.md --out 01_速通笔记_老奶版.docx
python "$B/build_docx.py" --src _source/notes_muscle.md  --out 01_速通笔记_薄肌版.docx
python "$B/build_docx.py" --src _source/questions.md     --out 02_通关题库_无答案.docx
python "$B/build_docx.py" --src _source/answers.md       --out 03_通关题库_答案解析.docx

python "$B/verify_docx.py" --src 01_速通笔记_老奶版.docx --out-dir _preview
```

逐页看图检查：封面、目录跳转、本章速览卡片不断裂、表格无孤行、四角水印不压字、无空白页。

源文件语法和排版参数见 [references/docx-spec.md](references/docx-spec.md)。

---

## 声音速查

| | 老奶版 | 薄肌版 |
| --- | --- | --- |
| 第一句 | 生活画面 | 能直接用的结论 |
| 句长 | ≤ 25 字为主 | 一行一条 |
| 术语 | 先解释再用 | 直接用，首次给 8 字定义 |
| 公式 | 每个字母读出来 | 公式 + 条件成对 |
| 幽默 | 温暖，不玩梗 | 可以吐槽题目 |
| 长度 | 每点 120–200 字 | 每点 40–80 字或 3 行表 |

**禁词（两版）**：显然、易证、不难发现、众所周知、宝子们。

---

## 自检清单

- [ ] 边界卡里的"排除范围"写清了砍掉什么、为什么。
- [ ] 每章有「本章速览」，建议用时之和与总预算一致。
- [ ] 全文没有 BOSS、血条、打怪、必杀技等游戏化字眼。
- [ ] 老奶版没有未解释的符号；薄肌版没有无条件的公式。
- [ ] 题库覆盖矩阵无空格，题目按考察点去重。
- [ ] 需要图的题都插了图，没有"如图"却无图。
- [ ] `check_docs.py` 全绿：题号对应、题干一致、无答案泄漏、图片存在、两版对齐、本章速览。
- [ ] 答案版题号与题库一一对应，每题都验算过。
- [ ] 每章一句话存在，「耍起」每章不超过一次。
- [ ] 四份 docx 都用 `verify_docx.py` 出过图，没有空白页。
- [ ] 文科记忆类：最后一页有「提词器」，且能单独使用。
- [ ] 材料没覆盖的地方都标了来源或 `（AI补充）`。

---

## 常见失误

| 失误 | 后果 | 纠正 |
| --- | --- | --- |
| 没定边界就开始写 | 写了半本不考的 | 回到 Phase 1 |
| 按 PPT 页数平分时间 | 重点章没时间 | 按优先级分，整章砍掉低优先级章 |
| 两版只是换语气 | 老奶版仍看不懂 | 老奶版重写结构，不是改词 |
| 答案版直接复用题库文件 | 无答案版被污染 | 两个独立源文件 |
| 题库和答案版各改各的题干 | 学生对着错误的题看解析 | 先定题库，再复制生成答案版；`check_docs.py` 会比对题干 |
| 秒杀法没写条件 | 考场套错 | 每招补"适用/不适用" |
| 题目只换数字当新题 | 题库注水 | `check_docs.py` 按归一化题干抓近重复 |
| 配图凭感觉找 | 图题对不上 | 先看 `图索引.md` 的页面文字，再 `--grid` 读坐标 |
| 表格跨页只剩一行 | 排版难看 | 渲染器已处理 ≤10 行表；更长的表拆开 |
| 没跑验收就交付 | 目录跳转失效/水印压字 | 必须出图看一眼 |
