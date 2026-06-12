# 北京高途 · 高中讲义转课件 Skill 库

本仓库存放北京高途高中各学科「讲义 docx + 参考PPT → 课件 pptx」生成 skill。

## 当前 Skill

| Skill | 学科 | 状态 |
|-------|------|------|
| [gaozhong-wuli-handout-to-courseware](./gaozhong-wuli-handout-to-courseware/) | 高中物理 | ✅ 已验证（几何光学 + 库仑定律双 case） |

## 快速开始

每个 skill 目录下有 `SKILL.md`，详细说明执行流程与注意事项。

```bash
# 示例：高中物理 几何光学
python3 gaozhong-wuli-handout-to-courseware/scripts/extract_handout.py "讲义.docx" -o output/第10讲/
python3 gaozhong-wuli-handout-to-courseware/scripts/build_pptx.py \
  --content output/第10讲/content.json \
  --ref "参考PPT.pptx" \
  -o output/第10讲/第10讲_课件.pptx
```
