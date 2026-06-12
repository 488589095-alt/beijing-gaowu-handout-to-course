# -*- coding: utf-8 -*-
"""
extract_handout.py — 北京高中物理「讲义 docx → content.json」解析器

物理讲义三大特性（决定本解析器与其它学科不同）：
  1. 公式 = OMML <m:oMath>（几何光学讲义 167 个）→ 按文档顺序编号存入 math_pool，
     正文里以 {"type":"math","id":"mNNN"} 段落片段就地引用（保持与文字交织的阅读顺序）。
  2. 图密集 = 光路图/电路图（讲义 28 张，多为 VML <v:imagedata>）→ 导出到 media/ 并按
     文档锚点挂到所在内容块（{"type":"img","ref":"imageN.ext"}）。
  3. 知识点驱动结构：模块(H1) → 考点(H2) → [知识精讲(H3) / 典型例题(H3)]；真题呈现(H1)。
     讲义【通常不含答案/解析】→ 标记 needs_answer，交 AI 生成（见 SKILL.md C 类）。

输出 output/<讲次>/content.json + media/。
用法: python3 extract_handout.py <讲义.docx> -o output/<讲次>/
"""
import argparse, json, re, shutil, zipfile
from pathlib import Path
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
ZWSP = "​"

import sys
sys.path.insert(0, str(Path(__file__).parent))
from omml import clean_omath, is_inline_simple, omath_to_unicode  # noqa: E402


def q(t, ns=W):
    return f"{{{ns}}}{t}"


def style_of(p):
    ppr = p.find(q("pPr"))
    if ppr is None:
        return ""
    s = ppr.find(q("pStyle"))
    return s.get(q("val")) if s is not None else ""


def load_rels(z):
    rels = {}
    try:
        x = etree.fromstring(z.read("word/_rels/document.xml.rels"))
        for r in x:
            rels[r.get("Id")] = r.get("Target")
    except KeyError:
        pass
    return rels


SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


def detect_pages(docx_path, out_dir):
    """把 docx 转 PDF，定位每个 例题N / 真题N 所在的 docx 页码（错了也认，按 docx 自身分页）。
       返回 {'例题1':3, ...}。失败返回 {}。"""
    import subprocess
    try:
        import fitz
    except Exception:
        return {}
    out_dir = Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    tmp = out_dir / "_handout.docx"
    shutil.copy(docx_path, tmp)
    try:
        subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf",
                        "--outdir", str(out_dir), str(tmp)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       timeout=180)
        pdf = out_dir / "_handout.pdf"
        d = fitz.open(str(pdf))
        pages = {}
        zhen = 0
        for i in range(d.page_count):
            t = d[i].get_text("text")
            for m in re.finditer(r"例题\s*(\d+)", t):
                k = f"例题{int(m.group(1))}"
                pages.setdefault(k, i + 1)
            for m in re.finditer(r"^\s*(\d+)[^\n]{0,4}(?:测量|如图|一|某)", t, re.M):
                pass
            if "真题" in t and not zhen:
                zhen = i + 1
        d.close()
        if zhen:
            pages["真题1"] = zhen
        for f in ("_handout.docx", "_handout.pdf"):
            (out_dir / f).unlink(missing_ok=True)
        return pages
    except Exception as e:
        print("⚠️ 页码检测失败(忽略):", e)
        return {}


class Extractor:
    def __init__(self, docx_path, out_dir):
        self.docx_path = docx_path
        self.z = zipfile.ZipFile(docx_path)
        self.out = Path(out_dir)
        self.media_dir = self.out / "media"
        self.rels = load_rels(self.z)
        self.math_pool = {}        # id -> cleaned oMath xml
        self._mi = 0
        self.media_used = {}       # imageN.ext -> exported path
        self._media_count = 0

    # ---- 公式：登记并返回 id ----
    def _reg_math(self, om):
        mid = f"m{self._mi:03d}"
        self._mi += 1
        self.math_pool[mid] = etree.tostring(clean_omath(om), encoding="unicode")
        return mid, om

    # ---- 图片：解析 rel → 导出 → 返回文件名 ----
    def _export_img(self, rid):
        target = self.rels.get(rid)
        if not target:
            return None
        src = "word/" + target.replace("../", "")
        try:
            data = self.z.read(src)
        except KeyError:
            return None
        ext = Path(target).suffix or ".png"
        name = f"image{self._media_count}{ext}"
        self._media_count += 1
        self.media_dir.mkdir(parents=True, exist_ok=True)
        (self.media_dir / name).write_bytes(data)
        self.media_used[name] = str(self.media_dir / name)
        return name

    # ---- 段落 → 有序 segments（text / math / img）----
    def segments(self, p):
        segs = []

        def flush_text(buf):
            t = "".join(buf).replace(ZWSP, "")
            if t:
                segs.append({"type": "text", "text": t})

        buf = []

        def walk(el):
            tag = el.tag
            if not isinstance(tag, str):
                return
            ln = etree.QName(el).localname
            if ln == "oMath" and tag.startswith(f"{{{M}}}"):
                flush_text(buf); buf.clear()
                mid, om = self._reg_math(el)
                if is_inline_simple(om):
                    u = omath_to_unicode(om)
                    segs.append({"type": "math", "id": mid, "inline": True, "unicode": u})
                else:
                    segs.append({"type": "math", "id": mid, "inline": False})
                return
            # 图片：DrawingML blip 或 VML imagedata
            if ln == "blip" and tag.startswith(f"{{{A}}}"):
                rid = el.get(q("embed", R)) or el.get(q("link", R))
                name = self._export_img(rid) if rid else None
                if name:
                    flush_text(buf); buf.clear()
                    segs.append({"type": "img", "ref": name})
                return
            if ln == "imagedata" and tag.startswith(f"{{{V}}}"):
                rid = el.get(q("id", R)) or el.get(q("href", R))
                name = self._export_img(rid) if rid else None
                if name:
                    flush_text(buf); buf.clear()
                    segs.append({"type": "img", "ref": name})
                return
            if ln == "t" and tag.startswith(f"{{{W}}}"):
                buf.append(el.text or "")
                return
            if ln == "tab" and tag.startswith(f"{{{W}}}"):
                buf.append("\t")
                return
            for c in el:
                walk(c)

        for c in p:
            walk(c)
        flush_text(buf)
        return segs

    def seg_text(self, segs):
        """segments 的纯文本投影（含 inline 公式的 unicode），用于判定/正则。"""
        out = []
        for s in segs:
            if s["type"] == "text":
                out.append(s["text"])
            elif s["type"] == "math" and s.get("inline"):
                out.append(s.get("unicode", ""))
            elif s["type"] == "math":
                out.append("〔公式〕")
        return "".join(out)

    def table_rows(self, tbl):
        rows = []
        for tr in tbl.findall(q("tr")):
            cells = []
            for tc in tr.findall(q("tc")):
                segs = []
                for p in tc.findall(q("p")):
                    segs += self.segments(p)
                cells.append(self.seg_text(segs).strip())
            rows.append(cells)
        return rows

    # ---- 主流程：线性遍历 body，按标题分组 ----
    def run(self):
        doc = etree.fromstring(self.z.read("word/document.xml"))
        body = doc.find(q("body"))
        blocks = []   # 线性块流：(kind, payload)
        for ch in body:
            ln = etree.QName(ch).localname
            if ln == "p":
                st = style_of(ch)
                segs = self.segments(ch)
                txt = self.seg_text(segs).strip()
                if not segs and not txt:
                    continue
                blocks.append(("p", {"style": st, "segs": segs, "text": txt}))
            elif ln == "tbl":
                blocks.append(("tbl", {"rows": self.table_rows(ch)}))

        return self.assemble(blocks)

    def assemble(self, blocks):
        # 标题 / 讲次号
        title_block = blocks[0][1]["text"] if blocks else ""
        mno = re.search(r"第[一二三四五六七八九十百\d]+讲", title_block)
        lecture_no = self._norm_lecture(mno.group(0)) if mno else "第00讲"
        title = re.sub(r"^第[一二三四五六七八九十百\d]+讲\s*", "", title_block)
        title = re.sub(r"【.*?】", "", title).strip() or title_block

        data = {
            "subject": "高中物理", "lecture_no": lecture_no, "title": title,
            "grade": "高二", "teacher": "主讲老师：",
            "textbook_link": None, "modules": [], "exam": [],
            "math_pool": self.math_pool, "media": list(self.media_used.keys()),
        }

        cur_mod = cur_sec = cur_h3 = None
        in_exam = False
        cur_ex = None  # 当前例题

        def new_section(name):
            return {"name": name, "knowledge": [], "examples": []}

        def push_example():
            nonlocal cur_ex
            if cur_ex is not None and cur_sec is not None:
                cur_sec["examples"].append(cur_ex)
            elif cur_ex is not None and in_exam:
                data["exam"].append(cur_ex)
            cur_ex = None

        for kind, pl in blocks[1:]:
            if kind == "tbl":
                rows = pl["rows"]
                if cur_h3 == "knowledge" and cur_sec is not None:
                    cur_sec["knowledge"].append({"type": "table", "rows": rows})
                elif data["textbook_link"] is None and cur_mod is None:
                    data["textbook_link"] = rows
                continue

            st, segs, txt = pl["style"], pl["segs"], pl["text"]

            # 标题层级
            if st == "1":   # H1 模块 / 教材链接 / 真题呈现
                push_example()
                if txt == "教材链接":
                    cur_mod = cur_sec = cur_h3 = None
                    continue
                # 真题/拓展类"做题区"H1（其下是编号题，无 例题N/知识精讲 H3）
                if any(k in txt for k in ("真题", "未来你会遇见", "拓展", "演练",
                                          "巩固提升", "强化训练", "课后")):
                    in_exam = True
                    data.setdefault("exam_section", txt)
                    cur_mod = cur_sec = cur_h3 = None
                    continue
                in_exam = False
                cur_mod = {"name": txt, "sections": []}
                data["modules"].append(cur_mod)
                cur_sec = None; cur_h3 = None
                continue
            if st == "2":   # H2 考点
                push_example()
                cur_sec = new_section(txt)
                if cur_mod is None:
                    cur_mod = {"name": txt, "sections": []}
                    data["modules"].append(cur_mod)
                cur_mod["sections"].append(cur_sec)
                cur_h3 = None
                continue
            if st == "3":   # H3 知识精讲 / 典型例题
                push_example()
                if "知识精讲" in txt:
                    cur_h3 = "knowledge"
                elif "典型例题" in txt or "例题" in txt:
                    cur_h3 = "examples"
                else:
                    cur_h3 = "knowledge"
                # H1 模块下直接接 H3（全反射型，无 H2）→ 建一个同名 section
                if cur_sec is None and cur_mod is not None:
                    cur_sec = new_section(cur_mod["name"])
                    cur_mod["sections"].append(cur_sec)
                continue

            # 正文段落
            proj = txt
            mex = re.match(r"^\s*例题\s*(\d+)", proj)
            mopt = re.match(r"^\s*([A-E])[．.。]", proj)
            mqno = re.match(r"^\s*(\d+)", proj)  # 真题编号（数字后直接接题干）

            if in_exam:
                if mqno and not mopt and (cur_ex is None or cur_ex.get("options")):
                    push_example()
                    cur_ex = {"no": f"真题{mqno.group(1)}", "stem_segs": segs,
                              "stem_imgs": [], "options": [], "answer": None,
                              "analysis": None, "needs_answer": True}
                    continue
            if cur_h3 == "examples":
                if mex:
                    push_example()
                    cur_ex = {"no": f"例题{mex.group(1)}", "stem_segs": segs,
                              "stem_imgs": [], "options": [], "answer": None,
                              "analysis": None, "needs_answer": True}
                    continue

            if cur_ex is not None and (cur_h3 == "examples" or in_exam):
                if mopt:
                    cur_ex["options"].append({"label": mopt.group(1), "segs": segs,
                                              "imgs": [s["ref"] for s in segs
                                                       if s["type"] == "img"]})
                    continue
                # 续接题干 / 图 / 多问
                imgs = [s["ref"] for s in segs if s["type"] == "img"]
                if imgs and not [s for s in segs if s["type"] != "img"]:
                    cur_ex["stem_imgs"] += imgs
                else:
                    # 多问 (1)(2) 或题干续行
                    cur_ex.setdefault("extra", []).append(segs)
                continue

            # 知识精讲正文块
            if cur_h3 == "knowledge" and cur_sec is not None:
                imgs = [s["ref"] for s in segs if s["type"] == "img"]
                non = [s for s in segs if s["type"] != "img"]
                if imgs and not non and cur_sec["knowledge"]:
                    # 纯图 → 挂到上一块
                    last = cur_sec["knowledge"][-1]
                    last.setdefault("imgs", []).extend(imgs)
                else:
                    cur_sec["knowledge"].append({"type": "para", "segs": segs,
                                                 "imgs": imgs})
                continue

        push_example()
        # 兜底：任何导出但未挂到任何块的图，挂到最后一道题/最后一个知识块，杜绝漏图
        def _all_refs():
            r = set()
            for m in data["modules"]:
                for sec in m["sections"]:
                    for b in sec["knowledge"]:
                        if b.get("type") == "para":
                            r |= {s["ref"] for s in b["segs"] if s["type"] == "img"}
                            r |= set(b.get("imgs", []))
                    for e in sec["examples"]:
                        r |= set(e.get("stem_imgs", []))
                        r |= {s["ref"] for s in e["stem_segs"] if s["type"] == "img"}
                        for o in e.get("options", []):
                            r |= set(o.get("imgs", []))
                            r |= {s["ref"] for s in o["segs"] if s["type"] == "img"}
                        for ext in e.get("extra", []):
                            r |= {s["ref"] for s in ext if s["type"] == "img"}
            for e in data["exam"]:
                r |= set(e.get("stem_imgs", []))
                r |= {s["ref"] for s in e["stem_segs"] if s["type"] == "img"}
            return r
        orphans = [im for im in self.media_used if im not in _all_refs()]
        if orphans:
            sink = data["exam"][-1] if data["exam"] else None
            if sink is None:
                for m in reversed(data["modules"]):
                    for sec in reversed(m["sections"]):
                        if sec["examples"]:
                            sink = sec["examples"][-1]; break
                    if sink:
                        break
            if sink is not None:
                sink.setdefault("stem_imgs", []).extend(orphans)
                print(f"⚠️ 兜底挂载孤儿图到 {sink.get('no','?')}: {orphans}")
            else:
                print(f"⚠️ 孤儿图无处挂载: {orphans}")
        # 题目页码（docx 自身分页，错了也认）
        pages = detect_pages(self.docx_path, self.out)
        for m in data["modules"]:
            for sec in m["sections"]:
                for ex in sec["examples"]:
                    ex["page"] = pages.get(ex["no"])
        for ex in data["exam"]:
            ex["page"] = pages.get(ex["no"])
        data["pages"] = pages
        # 统计
        data["_stats"] = {
            "modules": len(data["modules"]),
            "examples": sum(len(s["examples"]) for m in data["modules"]
                            for s in m["sections"]),
            "exam": len(data["exam"]),
            "formulas": len(self.math_pool),
            "images": len(self.media_used),
        }
        return data

    @staticmethod
    def _norm_lecture(s):
        cn = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7,
              "八": 8, "九": 9, "十": 10, "十一": 11, "十二": 12, "十三": 13,
              "十四": 14, "十五": 15, "十六": 16}
        m = re.search(r"第([一二三四五六七八九十百\d]+)讲", s)
        if not m:
            return s
        v = m.group(1)
        if v.isdigit():
            n = int(v)
        elif v in cn:
            n = cn[v]
        elif v.startswith("一十"):
            n = 10 + cn.get(v[2:], 0)
        elif v.endswith("十"):
            n = cn.get(v[0], 1) * 10
        else:
            n = cn.get(v, 0)
        return f"第{n:02d}讲"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    ex = Extractor(a.docx, out)
    data = ex.run()
    (out / "content.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"→ {out/'content.json'}")
    print("统计:", json.dumps(data["_stats"], ensure_ascii=False))


if __name__ == "__main__":
    main()
