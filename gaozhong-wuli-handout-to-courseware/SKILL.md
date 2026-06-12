---
name: gaozhong-wuli-handout-to-courseware
description: 北京高中物理「讲义 docx + 参考PPT → 课件 pptx」生成 skill（第10讲几何光学 case 跑通）。物理讲义的公式是 OMML 数学对象（非文字非图片）、且图密集（光路图/电路图）——本 skill 用 OMML 原生移植（a14 可编辑公式 + 图片兜底）+ 公式随锚点内联 + 光路图随锚点落页解决；无独立PPT模版时，参考PPT 同时充当蓝图+版式+设计token三重来源（复用其具名版式：首页/章节页/知识讲解/典型例题/真题呈现/结束）。讲义无答案→AI生成答案解析并标"待师审"+出审核清单。当用户提到"物理讲义转课件"、"高中物理出课件/做PPT"、"几何光学/电路/力学讲义转PPT"、"第N讲物理课件"时使用。注意与 handout-to-courseware（小学数学）、yuwen-（语文）、gaozhong-yingyu-（英语）、gaozhong-huaxue-（化学）区分：高中物理用这个。
---

# 北京高中物理 · 讲义 → 课件 PPT（v1）

```
讲义.docx ─extract_handout.py─▶ content.json（结构 + math_pool[OMML] + media[光路图]）─┐
AI 生成(C类:答案/解析/小结) ──▶ ai_scaffold.json + 《AI生成内容审核清单》──────────────┼─ build_pptx.py ─▶ 课件.pptx
参考PPT.pptx（具名版式=品牌底+页眉；蓝图；设计token）──────────────────────────────┘   + slide_structure.json
                                                                                          → verify.py 校对
```

## 物理三大特性（本 skill 核心能力，区别于其它学科）
1. **公式 = OMML `<m:oMath>`**（几何光学讲义 167 个、参考PPT 564 个），既非文字也非图片。
   → 默认 **native 模式**：块公式以**纯原生公式形状** `<p:sp>…<a14:m><m:oMathPara><m:oMath>`（无 mc/无图片）插入，
   **WPS/PowerPoint 渲染成可编辑『公式格式』**（`omml.math_native_xml`）；
   简单行内公式(θ₁、c>v)转 **Unicode** 进句子，且**单字母变量斜体/单位·数字·多字母函数正体**（`add_var_runs`，物理排版惯例）。
   `--formula hybrid` 可改"原生+图片兜底"（LibreOffice 等也能看，但公式变图片）。**讲义公式数 ↔ 成片 a14 数对账**。
2. **图密集**：光路图/电路图是内容核心 → **随讲义文档锚点内联**落到所在知识块/例题页（与英语"没图别插"相反）。
3. **无独立 PPT 模版 → 高保真靠"克隆参考PPT真页"**：参考PPT 同时是 蓝图+版式+fixture+token。
   **封面/模块标题/结束页 直接 `clone_slide` 克隆参考PPT真页再换文字**（保留圆角框/吉祥物/精确字体字号位置——
   勿用 token 重建，否则图形/字号对不上）；知识/例题页用具名版式 `add_slide` + 自加内容。最后删参考PPT原成品页。
   例题页右上加**难度标**(仿参考"P81►►►"区)：讲义无逐题页码→只标讲次难度档(【基础】=▶)，页码留老师补。
   **例题严格一页**：题干+题图+ABCD选项全在一页(题干字号自适应缩放保证不溢出，`r_example`)；
   **答案/解析不出现在页面里**→放**画布外右侧**(x>10in，放映看不到、编辑态老师可见，`_offcanvas_answer`)。
   课后测/课后练习/生活应用实景拓展/目录=讲义无的B类人工增量→**不生成**（见 output 同款《讲义vs参考PPT_差异对照》）。

## 三条铁律
1. **讲义有的 100% 转入**：公式用原生OMML、光路图张张落页、严格按参考PPT样式（微软雅黑+TNR，强调红#FF0000）。
2. **讲义没有、参考PPT有的人工内容**：课前/课后测题、配图插画、生活应用拓展实景页、页码标 → **不生成**；
   教学脚手架（**答案/解析**、知识总结小结）→ **AI 生成**，页面标"（AI 生成·待老师审核修改）"，并出《AI生成内容审核清单》。
3. **公式与图可追溯**：math_pool / media 每项带 id；自检对账数量。

## 结构骨架（知识点驱动，非题型分页）
```
封面 →[教材链接]→ 每模块[章节页(Part N) → 考点标题 → 知识精讲(概念+公式+光路图,流式分页)
                          → 每例题(题页→答页AI)] → 真题(题页→答页AI) → 知识总结(AI) → 结束
```
详见 [references/structure_recipe.md](references/structure_recipe.md) · [references/标杆蓝图_物理.md](references/标杆蓝图_物理.md)

## 执行流程
```bash
# 0.【Gate·三拆解】参考PPT → 三件套（新讲次/新参考PPT必跑；已沉淀几何光学版可直接复用）
python3 {SKILL_DIR}/scripts/dissect_template.py  "<参考PPT.pptx>" -o {SKILL_DIR}/references/   # 版式/标题槽
python3 {SKILL_DIR}/scripts/extract_styles.py    "<参考PPT.pptx>" --slides 1,2,3,8,48,57 -o references/style_raw.json  # 逐run样式→curate style_form
python3 {SKILL_DIR}/scripts/blueprint_benchmark.py "<参考PPT.pptx>" -o references/标杆蓝图_raw.txt  # 逐页蓝图→curate 标杆蓝图_物理.md
#   ★给到的参考PPT都要拆蓝图★；确认具名版式名(首页/章节页/知识讲解/典型例题/真题呈现/4_自定义版式)

# 1. 解析讲义（公式→math_pool / 光路图→media+锚点 / 结构分块）
python3 {SKILL_DIR}/scripts/extract_handout.py "<讲义.docx>" -o output/<讲次>/
#   产出 content.json（含 _stats：modules/examples/exam/formulas/images）+ media/

# 2.【人工/AI】生成 C 类 → output/<讲次>/ai_scaffold.json（schema 见 references/content_schema.md）
#   ★讲义通常不含答案/解析★ → AI 逐题解物理题写 {answer, analysis}（标 _confidence），写 summary 小结
#   ⚠️ 全是草稿：逐项列入《AI生成内容审核清单.md》交老师审；答案正确性老师必核
#   （记忆铁律：AI 内容必须配带解析+参考答案的 answer key 供老师审核）

# 3. 渲染（克隆参考PPT封面/模块/结束真页 + 原生公式 + 光路图内联 + 题页→答页 + 答案红 + 难度标）
python3 {SKILL_DIR}/scripts/build_pptx.py \
  --content output/<讲次>/content.json --ref "<参考PPT.pptx>" \
  -o "output/<讲次>/<讲次>_课件.pptx"          # 默认 --formula native（公式格式,WPS/PPT可编辑）
#   公式默认原生(WPS/PowerPoint可编辑·可正确斜体)；如需在 LibreOffice 等看公式加 --formula hybrid
#   ⚠️ native 公式在 LibreOffice/预览/Google 里不显示 → 校对务必用 WPS/PowerPoint 打开

# 4. 校对（对照讲义抽取 + 蓝图）
python3 {SKILL_DIR}/scripts/verify.py --content output/<讲次>/content.json --pptx "<课件>"
#   检查：页数/角色页计数 · 公式守恒(a14≥块公式) · 图落页(嵌图≥讲义图) · AI待审标注 · 无参考PPT残留
#   本机可视校对：soffice --convert-to pdf 课件 → PyMuPDF 截图逐页看（LibreOffice 走公式兜底图，PowerPoint走原生）
```

## 人工确认 Gate
1. **slide_structure.json / 标杆蓝图_物理.md**（页序+角色）→ 老师确认结构
2. **《AI生成内容审核清单》**（所有 C 类，**尤其每道例题/真题的答案与解析**）→ 老师审核修改
3. 成片在 **PowerPoint/WPS** 逐页核对：公式原生渲染、换行/溢出、光路图位置（本机仅 LibreOffice 时看兜底图）

## 已验证用例
| 讲次 | 类型 | 产出 | 验证 |
|---|---|---|---|
| 第10讲 几何光学 | 知识精讲+14例题+1真题 | 54页 | 对标参考PPT(57页)；公式34 a14≥29块/光路图26张全落页；例题1解析√15/2已验 |

## 经验坑（必读）
- **保真度第一坑：高保真靠克隆参考PPT真页，不要 token 重建**。曾用 add_textbox 按拆解 token 重建封面→
  丢了圆角框(Shape2/3)/吉祥物/精确字号，"字体字号图形都对不上"。改 `clone_slide` 克隆封面/模块/结束真页+换文字后一致。
- **公式优先 native（公式格式）**：`math_native_xml` 出纯 `<a14:m><m:oMathPara><m:oMath>`（无 mc/无图）→ WPS/PowerPoint
  可编辑公式、且数学区自动正确斜体；裸 a14 在 LibreOffice 不显示(只 WPS/PPT)，故 LibreOffice 校对要 `--formula hybrid`。
  教训：曾用 mc+图片兜底，WPS 取了图片兜底→"公式变图片不可编辑"，故默认改 native。
- **正斜体**(物理惯例)：行内公式单字母变量(n,c,θ)斜体、单位(cm,C)·数字·多字母函数(sin)正体 → `add_var_runs` 逐段判定；
  块公式走原生 OMML 数学区自动处理。
- **必删继承占位符**：`add_slide(版式)` 会带入版式的空 OBJECT 占位符→显示"单击此处添加文本"/图片图标。
  `strip_placeholders` 删之。版式真规格(典型例题 idx10 正文=20pt·idx11 在画布外=答案区)用作字号/位置参考。
- **例题严格一页·题图选项相邻**：题干字号从 20pt 起按 题干+图+引导+选项 总高自适应下调；**无图不留空**；
  选项行高按**实际折行数**算(长选项 2 行→别按 1 行排否则叠印)；图形选项横排一行。
- **图可能藏在题干中段/extra**：如例题5 棱镜图+数据表在"题干→[图]→根据…下列说法"之间（extra 段含 img）→
  `_example_figs_and_leadin` 把 stem_imgs+extra+stem_segs 的图都收出来按序放，引导文字接图后。
- **右上角标(仿参考)**：讲义页码框(由 `detect_pages` 把 docx 转 PDF 定位每题页→P3/P6…) + 难度旗帜(`harvest_flag` 挪用参考PPT画布外真旗帜图)；不放左上例题标签。
- **完整性校对(防漏图漏内容)**：build 末尾 `_coverage_check` 对账「讲义图数→成片嵌图数」「块公式≥→成片a14」，漏则醒目告警。两类已修的漏：
  ① 图形选项(例题6/7 每项一图)——`_split_options` 拆分时务必把 img 归到 `imgs`(否则被当文字选项,`fill_para` 跳过图→丢)；
  ② 真题/拓展类 H1(如"未来你会遇见")下的编号题——exam-section 识别要含 真题/未来你会遇见/拓展/演练/巩固/课后，否则整题(文字+图)被丢；
  并加 extract 末尾"孤儿图兜底挂载到最后一题"。**新讲次跑完务必看 _coverage_check 全 ✅**。
- **清洗 OMML**：去 U+200B(零宽空格,渲染成□缺字框) + 去 `w:` 命名空间属性(PPT无效) + 根式 degHide 缺 `<m:deg>` 补空 deg(否则兜底图 √□)。见 `omml.clean_omath`。
- **公式被拆**：讲义里 θ₁ 常拆成相邻两个 oMath(θ / 空底下标1)→ 行内 Unicode 自然拼接为 θ₁；故行内走 Unicode、块公式走形状。
- **画布 10×7.5 in**（非英语16.67×12.5）；所有坐标按此，换模版重测。
- **具名版式自带标签**：典型例题版式已有居中"典型例题"字 → 别再加同名标题(会重叠)，只在左上放 例题号·考点。
- **讲义全无答案**：物理讲义只给题不给解 → 答案/解析 100% 是 AI 增量，务必出审核清单、老师核对正确性。
- **题图过大挤掉选项**：题图 max 5.2×2.5in，保证 4 选项同页(否则 D 项溢出成孤页)。
- **中文行高估算必须按全角**(PATCH-001)：Flow `_est_lines` 逐字累宽——中文/全角标点=1.0em、西文=0.55em；
  用西文系数会严重低估中文行数→文本框过矮→长句换行后上下**叠印**(知识点⑤⑥⑦最典型)。
  另：所有文本框 `auto_size=NONE` 关 spAutoFit(框不自动长高去压下一框)。验证：见 verify 后跑「宽框×宽框竖直相交」检测应=0。
- 图是 VML `<v:imagedata>`(非 DrawingML blip)→ extract 两种都抓。
- 课前/课后测/练习、生活应用实景拓展页=人工增量(外部题源/实景图)，不生成。

## 设计 token / schema / 骨架明细
[references/design_tokens.md](references/design_tokens.md) · [references/content_schema.md](references/content_schema.md) ·
[references/structure_recipe.md](references/structure_recipe.md) · [references/style_form.md](references/style_form.md)
