# content.json / ai_scaffold.json Schema（北京高中物理）

## content.json（extract_handout.py 产出）
```jsonc
{
  "subject": "高中物理", "lecture_no": "第10讲", "title": "几何光学",
  "grade": "高二", "teacher": "主讲老师：",
  "textbook_link": [["表头…"], …] | null,        // 教材链接表
  "math_pool": {                                   // ★所有公式：id → 干净 <m:oMath> XML
    "m000": "<m:oMath xmlns:m=…>…</m:oMath>", …    //   按文档顺序编号；clean_omath 已去U+200B/w:/补根式deg
  },
  "media": ["image0.jpg", …],                      // 导出到同目录 media/ 的图（光路图）
  "modules": [{
    "name": "光的折射",                            // 讲义 H1
    "sections": [{
      "name": "光的折射",                          // 讲义 H2（无H2则同模块名）
      "knowledge": [                               // 知识精讲块（有序）
        {"type":"para", "segs":[…], "imgs":["imageN.jpg"]},
        {"type":"table", "rows":[[…]]}
      ],
      "examples": [{                               // 典型例题
        "no":"例题1",
        "stem_segs":[…],                           // 题干 segments
        "stem_imgs":["image1.jpg"],                // 题图
        "options":[{"label":"A","segs":[…],"imgs":[]}],
        "extra":[[…]],                             // 多问(1)(2)/题干续行
        "answer": null, "analysis": null,          // ★讲义无→ null，待 AI 补
        "needs_answer": true
      }]
    }]
  }],
  "exam": [ {同 example 结构, "no":"真题1"} ],      // 真题呈现
  "_stats": {"modules":…, "examples":…, "exam":…, "formulas":…, "images":…}
}
```

### segments（正文有序片段，文字/公式/图交织）
```jsonc
[
  {"type":"text", "text":"折射率公式："},
  {"type":"math", "id":"m000", "inline":false},                 // 块公式→独立形状(a14+图)
  {"type":"math", "id":"m013", "inline":true, "unicode":"θ₁"},  // 简单行内→Unicode进句子
  {"type":"img",  "ref":"image0.jpg"}                           // 光路图
]
```
- `inline:true` ⇒ 渲染时用 `unicode` 直接作文字；`inline:false` ⇒ 取 `math_pool[id]` 出独立公式形状。
- 判定：`omml.is_inline_simple`（无分式/根式/积分/矩阵/函数 → 可行内）。

## ai_scaffold.json（C类·AI生成，须配《AI生成内容审核清单.md》）
```jsonc
{
  "_meta": "全部待老师审核修改",
  "例题1": {"answer":"C", "analysis":"…完整解析…", "_confidence":"high|draft|reviewed"},
  "例题2": {"answer":"（待老师核定）", "analysis":"思路：…", "_confidence":"draft"},
  …每题…,
  "真题1": {"answer":…, "analysis":…},
  "summary": ["1. 两个定律记心间", "2. n=sinθ₁/sinθ₂=c/v", …]   // 知识总结小结
}
```
键 = content 里的 `example.no` / `exam.no`；build 用其填答案页。无此文件则答案页留空待补。

## 解析容错（坑）
- 公式被拆成多个相邻 oMath（如 θ 与 ₁ 分属两个 oMath）→ 行内 Unicode 自然拼接为 θ₁。
- 选项可为：纯文字 / 纯公式（块）/ 每项一图（图形选项，如例题6/7）→ option 渲染分别处理。
- 根式 `<m:rad>` degHide 缺 `<m:deg>` → clean_omath 补空 deg（否则兜底图渲染成 √□）。
- 真题题干以「数字+中文」开头（"1测量…"）→ 编号正则不要求其后有空格/句点。
