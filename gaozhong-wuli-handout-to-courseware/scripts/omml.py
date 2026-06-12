# -*- coding: utf-8 -*-
"""
omml.py — 物理课件的「OMML 公式」核心模块（本 skill 的关键能力）

物理讲义里公式不是文字也不是图片，而是 OMML 数学对象 <m:oMath>（几何光学讲义 167 个）。
参考PPT 里也是原生 OMML（564 个），每个公式是一个独立形状，用
  <mc:AlternateContent><mc:Choice Requires="a14"> 原生公式 </mc:Choice>
                       <mc:Fallback> 公式图片 </mc:Fallback></mc:AlternateContent>
包装：PowerPoint/WPS 渲染可编辑的原生公式；其它阅读器(含 LibreOffice)渲染图片兜底。
本模块复刻这套机制：
  · clean_omath   : 去 U+200B(零宽空格,渲染成缺字框) + 去 w: 命名空间属性(PPT 里无效)
  · omath_to_unicode : 尽力把简单行内公式转 Unicode 文本(θ₁, n², a/b…)，让句子可读
  · is_inline_simple : 判定是否"简单行内"(无分式/根式/积分→可转 Unicode)
  · math_shape_xml : 生成 a14 原生 + 图片兜底 的独立公式形状 XML
  · render_math_images : 批量把公式经 docx→LibreOffice→PDF 渲染成 PNG(兜底图+本机校对)
"""
import copy, html, subprocess, sys
from pathlib import Path
from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
ZWSP = "​"
SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

_SUB = str.maketrans("0123456789+-=()aeoxhklmnpst", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓₕₖₗₘₙₚₛₜ")
_SUP = str.maketrans("0123456789+-=()n", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ")


def _q(tag, ns=M):
    return f"{{{ns}}}{tag}"


def clean_omath(om):
    """深拷贝 + 去 U+200B + 去 w: 命名空间元素 + 规整根式。返回干净的 <m:oMath> 元素。"""
    om = copy.deepcopy(om)
    for el in list(om.iter()):
        if not isinstance(el.tag, str):
            continue
        if el.tag.startswith(f"{{{W}}}"):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
            continue
        if el.text and ZWSP in el.text:
            el.text = el.text.replace(ZWSP, "")
    # 根式规整：degHide 开启但缺 <m:deg> 时补一个空 deg
    # （否则 LibreOffice 把被开方数渲染进隐藏的指数槽 → 显示空框 √□）
    for rad in om.findall(f".//{_q('rad')}"):
        if rad.find(_q("deg")) is None:
            e = rad.find(_q("e"))
            deg = etree.Element(_q("deg"))
            if e is not None:
                e.addprevious(deg)
            else:
                rad.append(deg)
    return om


def _text_of(el):
    return "".join(t.text or "" for t in el.findall(f".//{_q('t')}")).replace(ZWSP, "")


def is_inline_simple(om):
    """无分式 f / 根式 rad / 积分nary / 矩阵 m / 函数 / 上下极限 → 可转 Unicode 行内。"""
    for tag in ("f", "rad", "nary", "m", "func", "limLow", "limUpp", "groupChr", "bar"):
        if om.find(f".//{_q(tag)}") is not None:
            return False
    return True


def _to(s, table):
    """逐字符容错转上/下标，转不了的原样保留。"""
    return "".join(c.translate(table) for c in s)


def omath_to_unicode(om):
    """尽力把 OMML 转成 Unicode 行内文本（仅用于简单公式；复杂的走图片）。"""
    def _kids(el):
        return "".join(render(c) for c in el) if el is not None else ""

    def render(el):
        tag = etree.QName(el).localname if isinstance(el.tag, str) else ""
        if tag == "t":
            return (el.text or "").replace(ZWSP, "")
        if tag == "sSub":
            return _kids(el.find(_q("e"))) + _to(_kids(el.find(_q("sub"))), _SUB)
        if tag == "sSup":
            return _kids(el.find(_q("e"))) + _to(_kids(el.find(_q("sup"))), _SUP)
        if tag == "sSubSup":
            return (_kids(el.find(_q("e"))) + _to(_kids(el.find(_q("sub"))), _SUB)
                    + _to(_kids(el.find(_q("sup"))), _SUP))
        if tag == "d":
            return "(" + _kids(el.find(_q("e"))) + ")"
        return _kids(el)

    return "".join(render(c) for c in om).replace(ZWSP, "").strip()


# ════════ 行内原生公式（嵌进正文段落 <a:p>，不另起形状，避免偏移）════════
def math_inline_xml(om):
    """返回 <a14:m><m:oMath>…</m:oMath></a14:m> 串，作为 run 级元素插进段落。
       WPS/PowerPoint 行内渲染、自动斜体；放在文字之间不偏移。"""
    inner = etree.tostring(clean_omath(om), encoding="unicode")
    a = inner.find(">") + 1
    b = inner.rfind("</")
    return (f'<a14:m xmlns:a14="{A14}"><m:oMath xmlns:m="{M}">'
            f'{inner[a:b]}</m:oMath></a14:m>')


# ════════ 公式形状 ════════
def math_native_xml(om, sid, left, top, w, h):
    """纯原生公式形状（无 mc/无图片兜底）：PowerPoint/WPS 渲染成『公式格式』可编辑对象。
       结构同参考PPT：<p:sp>…<a:p><a14:m><m:oMathPara><m:oMath>…。
       注意：LibreOffice/Google 不支持 a14 → 这些阅读器看不到公式（目标是 WPS/PowerPoint）。"""
    L, T, Wd, Hd = (int(round(v * 914400)) for v in (left, top, w, h))
    inner = etree.tostring(clean_omath(om), encoding="unicode")
    a = inner.find(">") + 1
    b = inner.rfind("</")
    inner_body = inner[a:b]
    return (
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}" xmlns:a14="{A14}">'
        f'<p:nvSpPr><p:cNvPr id="{sid}" name="math{sid}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{L}" y="{T}"/><a:ext cx="{Wd}" cy="{Hd}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/>'
        f'<a:p><a:pPr/><a14:m><m:oMathPara xmlns:m="{M}">'
        f'<m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>'
        f'<m:oMath xmlns:m="{M}">{inner_body}</m:oMath></m:oMathPara></a14:m></a:p>'
        f'</p:txBody></p:sp>')


def math_shape_xml(om, sid, left, top, w, h, img_rel_id=None, fallback_text="[公式]"):
    """a14 原生 + 兜底（图片/文字）的 mc:AlternateContent 公式形状（hybrid 模式）。"""
    L, T, Wd, Hd = (int(round(v * 914400)) for v in (left, top, w, h))
    inner = etree.tostring(clean_omath(om), encoding="unicode")
    a = inner.find(">") + 1
    b = inner.rfind("</")
    inner_body = inner[a:b]   # <m:oMath> 的内部内容

    if img_rel_id:
        fb = (f'<p:pic xmlns:p="{P}" xmlns:a="{A}" '
              f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
              f'<p:nvPicPr><p:cNvPr id="{sid}" name="mathimg{sid}"/>'
              f'<p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
              f'<p:blipFill><a:blip r:embed="{img_rel_id}"/>'
              f'<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
              f'<p:spPr><a:xfrm><a:off x="{L}" y="{T}"/><a:ext cx="{Wd}" cy="{Hd}"/></a:xfrm>'
              f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    else:
        fb = (f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{sid}" name="mathfb{sid}"/>'
              f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
              f'<a:xfrm><a:off x="{L}" y="{T}"/><a:ext cx="{Wd}" cy="{Hd}"/></a:xfrm>'
              f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/>'
              f'<a:p><a:r><a:rPr lang="zh-CN"/><a:t>{html.escape(fallback_text)}</a:t></a:r></a:p>'
              f'</p:txBody></p:sp>')

    return (
        f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:a14="{A14}">'
        f'<mc:Choice Requires="a14">'
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr>'
        f'<p:cNvPr id="{sid}" name="math{sid}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{L}" y="{T}"/><a:ext cx="{Wd}" cy="{Hd}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        f'<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>'
        f'<a:p><a:pPr/><a14:m><m:oMathPara xmlns:m="{M}"><m:oMath xmlns:m="{M}">'
        f'{inner_body}</m:oMath></m:oMathPara></a14:m></a:p></p:txBody></p:sp>'
        f'</mc:Choice>'
        f'<mc:Fallback>{fb}</mc:Fallback>'
        f'</mc:AlternateContent>')


# ════════ 批量把公式渲染成 PNG（docx→LibreOffice→PDF→trim）════════
def render_math_images(math_pool, out_dir, dpi=200):
    """math_pool={id: oMath_xml_string 或 元素}。返回 {id: png_path}。
       每个公式占 docx 一页 → PDF 每页一公式 → 裁白边存 png。"""
    import fitz
    from PIL import Image, ImageChops
    from docx import Document
    out_dir = Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    ids = list(math_pool.keys())
    if not ids:
        return {}
    d = Document()
    for i, mid in enumerate(ids):
        v = math_pool[mid]
        om = etree.fromstring(v) if isinstance(v, str) else v
        p = d.add_paragraph()
        p._p.append(clean_omath(om))
        if i != len(ids) - 1:
            d.add_page_break()
    tmp_docx = out_dir / "_formulas.docx"
    d.save(str(tmp_docx))
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir",
                    str(out_dir), str(tmp_docx)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = out_dir / "_formulas.pdf"
    doc = fitz.open(str(pdf))
    res = {}
    for i, mid in enumerate(ids):
        if i >= doc.page_count:
            break
        pix = doc[i].get_pixmap(dpi=dpi)
        png = out_dir / f"math_{mid}.png"
        pix.save(str(png))
        im = Image.open(png).convert("RGB")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bbox = ImageChops.difference(im, bg).getbbox()
        if bbox:
            pad = 6
            l, t, r, b = bbox
            im = im.crop((max(0, l - pad), max(0, t - pad),
                          min(im.width, r + pad), min(im.height, b + pad)))
            im.save(str(png))
        res[mid] = str(png)
    doc.close()
    return res


if __name__ == "__main__":
    import zipfile, json
    path = sys.argv[1]
    mode = sys.argv[2] if len(sys.argv) > 2 else "--unicode"
    z = zipfile.ZipFile(path)
    doc = etree.fromstring(z.read("word/document.xml"))
    omaths = doc.findall(f".//{_q('oMath')}")
    print(f"oMath 总数: {len(omaths)}")
    if mode == "--unicode":
        simple = 0
        for i, om in enumerate(omaths):
            s = is_inline_simple(om)
            simple += s
            if i < 30:
                u = omath_to_unicode(om) if s else "〔复杂→图片〕"
                print(f"  {i:3d} {'inline' if s else 'block '} | {u[:50]!r}")
        print(f"简单行内可转Unicode: {simple}/{len(omaths)}")
    elif mode == "--images":
        pool = {f"m{i:03d}": etree.tostring(om, encoding="unicode")
                for i, om in enumerate(omaths[:8])}
        res = render_math_images(pool, "/tmp/math_imgs_test")
        print("rendered:", json.dumps(res, ensure_ascii=False, indent=1))
