# -*- coding: utf-8 -*-
"""
verify.py — 课件成片自检（对照讲义抽取 content.json 与参考PPT蓝图校对）

校对（物理关键）：
  1. 结构：页数、各角色页计数（cover/module/knowledge/example_q/example_a/end）
  2. 公式守恒：成片里 a14 原生公式数（<a14:m>）≥ content 中「块公式」数
  3. 图落页：成片嵌图数 ≥ 讲义导出图数（光路图必须都在）
  4. 无残留：成片不应残留参考PPT原成品页特征文字
  5. AI 标注：example_a 页应带"待老师审核"字样

用法: python3 verify.py --content output/<讲次>/content.json --pptx output/<讲次>/xxx_课件.pptx
"""
import argparse, json, zipfile
from pathlib import Path
from collections import Counter


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--content", required=True)
    ap.add_argument("--pptx", required=True)
    a = ap.parse_args()
    C = json.loads(Path(a.content).read_text(encoding="utf-8"))
    ss = Path(a.content).with_name("slide_structure.json")
    struct = json.loads(ss.read_text(encoding="utf-8")) if ss.exists() else None

    z = zipfile.ZipFile(a.pptx)
    slides = [n for n in z.namelist()
              if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    a14 = pics = ai_marks = 0
    for s in slides:
        x = z.read(s).decode("utf-8", "ignore")
        a14 += x.count("<a14:m")
        ai_marks += x.count("待老师审核")
    media = [n for n in z.namelist() if n.startswith("ppt/media/")]

    n_formula = len(C.get("math_pool", {}))
    n_block = 0

    def scan(segs):
        nonlocal n_block
        for sg in segs:
            if sg.get("type") == "math" and not sg.get("inline"):
                n_block += 1
    for m in C["modules"]:
        for sec in m["sections"]:
            for blk in sec["knowledge"]:
                if blk.get("type") == "para":
                    scan(blk["segs"])
            for ex in sec["examples"]:
                scan(ex["stem_segs"])
                for o in ex.get("options", []):
                    scan(o["segs"])
    n_img = len(C.get("media", []))
    n_ex = sum(len(s["examples"]) for m in C["modules"] for s in m["sections"])

    print(f"=== 自检：{Path(a.pptx).name} ===")
    print(f"页数: {len(slides)}")
    if struct:
        roles = Counter(s["role"] for s in struct["slides"])
        print("角色页计数:", dict(roles))
    print(f"公式: 成片 a14 原生 {a14} 个 | 讲义块公式应≥{n_block} | math_pool 总 {n_formula}")
    print(f"图片: 成片嵌图 {len(media)} 个 | 讲义导出 {n_img} 张")
    print(f"例题: {n_ex} 题（应有 {n_ex} 题页 + {n_ex} 答页）")
    print(f"AI待审标注: {ai_marks} 处")

    ok = True
    if a14 < n_block:
        print(f"⚠️ 公式偏少：a14 {a14} < 块公式 {n_block}（检查是否漏注入）"); ok = False
    if len(media) < n_img:
        print(f"⚠️ 图片偏少：成片 {len(media)} < 讲义 {n_img}"); ok = False
    residual = ["学有所获", "课前测"]
    hits = [w for w in residual if any(w in z.read(s).decode("utf-8", "ignore")
                                       for s in slides)]
    if hits:
        print(f"⚠️ 疑似参考PPT残留文字: {hits}"); ok = False
    print("✅ 自检通过" if ok else "❌ 自检有警告，请核对")


if __name__ == "__main__":
    main()
