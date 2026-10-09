# -*- coding: utf-8 -*-
"""
页眉页脚与分节页码（thesis-format-conformer 的第四步，可选）

把「已套用规范样式」的论文按三段式重建分节，并建立页码页脚与奇偶页页眉：

    第 1 节  封面            —— 无页眉、无页脚（不显示页码）
    第 2 节  摘要/Abstract/目录 —— 无页眉；页脚居中 PAGE 域，大写罗马数字，从 I 起
    第 3 节  正文起（含参考文献/致谢） —— 奇数页页眉 = STYLEREF「标题 1」(章号+章名)，
                                    偶数页页眉 = 论文题目；页脚居中 PAGE 域，阿拉伯数字，从 1 起

设计要点（照 2026-10-08 对规范论文与学生最终版的实测）：
  * 分节靠「段落内嵌 sectPr」实现：某段 pPr 里的 sectPr 描述"到此为止"这一节。
  * 页眉页脚按节挂 headerReference / footerReference；未定义的节继承上一节。
  * 奇偶页不同需在 settings.xml 打开 <w:evenAndOddHeaders/>。
  * 章标题的「第X章」由段落级自动编号产生，故 STYLEREF 用 \\n 开关取编号、再取一次取文本。
  * settings.xml 打开 <w:updateFields/>，Word 打开时自动刷新域（否则页码/章名为缓存值）。

用法：
    python fmt_hf.py <已规范化.docx> -o <输出.docx> --title "论文题目"
"""
import argparse
import os
import re
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx_kit as kit
import spec_rules as spec

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

HDR_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FTR_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"


def w(tag):
    return "{%s}%s" % (W_NS, tag)


def r(tag):
    return "{%s}%s" % (R_NS, tag)


NSDECL = 'xmlns:w="%s" xmlns:r="%s"' % (W_NS, R_NS)


def _field(instr, cached=""):
    """构造一个域（begin / instrText / separate / 缓存结果 / end）。"""
    cache = ""
    if cached:
        cache = "<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>" % _esc(cached)
    return (
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
        "<w:r><w:instrText xml:space=\"preserve\">%s</w:instrText></w:r>"
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        "%s"
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
    ) % (_esc(instr), cache)


def _esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def make_header_styleref(hdr_style_id, chapter_cached):
    """奇数页页眉：STYLEREF 取章号 + 章名（两个域拼接）。"""
    f1 = _field(' STYLEREF  "%s" \\n  \\* MERGEFORMAT ' % spec.HEADER_STYLEREF,
                chapter_cached.split("  ")[0] if chapter_cached else "")
    f2 = _field(' STYLEREF  "%s"  \\* MERGEFORMAT ' % spec.HEADER_STYLEREF,
                "  ".join(chapter_cached.split("  ")[1:]) if chapter_cached else "")
    xml = (
        '<w:hdr %s><w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>'
        "%s"
        '<w:r><w:t xml:space="preserve">  </w:t></w:r>'
        "%s"
        "</w:p></w:hdr>"
    ) % (NSDECL, hdr_style_id, f1, f2)
    return xml


def make_header_text(hdr_style_id, text):
    """偶数页页眉：直接显示论文题目。"""
    return (
        '<w:hdr %s><w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
        "</w:p></w:hdr>"
    ) % (NSDECL, hdr_style_id, _esc(text))


def make_footer_page(ftr_style_id, cached="1"):
    """页脚：居中 PAGE 域。"""
    return (
        '<w:ftr %s><w:p><w:pPr><w:pStyle w:val="%s"/><w:jc w:val="center"/></w:pPr>'
        "%s"
        "</w:p></w:ftr>"
    ) % (NSDECL, ftr_style_id, _field("PAGE   \\* MERGEFORMAT", cached))


# ---------------------------------------------------------------------------
# 分节定位
# ---------------------------------------------------------------------------
def find_boundaries(blocks, styles):
    """
    返回 (front_no, body_no)：
      front_no = 第一个前置篇级标题（摘要/Abstract/目录/声明）的块号
      body_no  = 第一个章标题的块号
    找不到则对应值为 None。
    """
    front_no = None
    body_no = None
    for kind, el, no in blocks:
        if kind != "p":
            continue
        text = kit.para_text(el)
        props, sid = kit.para_effective(el, styles)
        role = kit.detect_role(text, props, styles, sid)
        compact = text.strip().replace(" ", "").replace("\u3000", "").lower()
        if front_no is None and role in ("major_title", "cover_title"):
            if compact in [t.lower() for t in spec.HF_FRONT_TITLES]:
                front_no = no
        if body_no is None and role == "chapter_title":
            body_no = no
        if front_no is not None and body_no is not None:
            break
    return front_no, body_no


def first_chapter_text(blocks, styles, body_no):
    if body_no is None:
        return ""
    for kind, el, no in blocks:
        if kind == "p" and no == body_no:
            return kit.para_text(el).strip()
    return ""


# ---------------------------------------------------------------------------
# 部件与关系
# ---------------------------------------------------------------------------
def next_free_id(existing, prefix, start=1):
    i = start
    while ("%s%d.xml" % (prefix, i)) in existing:
        i += 1
    return i


def add_part(parts, rels_root, ct_root, kind, xml_text, used_names):
    """把 header/footer 部件写入 parts，并登记关系与内容类型。返回 rId。"""
    prefix = "header" if kind == "header" else "footer"
    idx = next_free_id(used_names, prefix)
    name = "%s%d.xml" % (prefix, idx)
    used_names.add(name)
    parts["word/" + name] = xml_text.encode("utf-8")
    # rId
    existing_ids = {rel.get("Id") for rel in rels_root}
    rid = "rId%d" % (max([int(x[3:]) for x in existing_ids if re.fullmatch(r"rId\d+", x)] or [0]) + 1)
    rel = etree.SubElement(rels_root, "{%s}Relationship" % REL_NS)
    rel.set("Id", rid)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + prefix)
    rel.set("Target", name)
    # 内容类型
    pn = "/word/" + name
    if not any(o.get("PartName") == pn for o in ct_root):
        ov = etree.SubElement(ct_root, "{%s}Override" % CT_NS)
        ov.set("PartName", pn)
        ov.set("ContentType", HDR_CT if kind == "header" else FTR_CT)
    return rid


def _first(root, *tags):
    """返回第一个存在的子元素（注意：lxml 元素的真值判断不可靠，必须用 is not None）。"""
    for t in tags:
        el = root.find(w(t))
        if el is not None:
            return el
    return None


def insert_even_odd_headers(settings_root):
    """
    在 settings.xml 打开 evenAndOddHeaders。
    按 CT_Settings 顺序，它紧随 defaultTabStop / autoHyphenation 之后。
    """
    if settings_root.find(w("evenAndOddHeaders")) is not None:
        return False
    el = etree.Element(w("evenAndOddHeaders"))
    anchor = _first(settings_root, "autoHyphenation", "defaultTabStop", "zoom")
    if anchor is not None:
        anchor.addnext(el)
    else:
        settings_root.insert(0, el)
    return True


def ensure_update_fields(settings_root):
    """打开 updateFields，使 Word 打开时刷新域（页码/章名不再是缓存值）。"""
    if settings_root.find(w("updateFields")) is not None:
        return False
    el = etree.Element(w("updateFields"))
    el.set(w("val"), "true")
    anchor = _first(settings_root, "showXMLTags", "alwaysMergeEmptyNamespace",
                    "hdrShapeDefaults", "footnotePr")
    if anchor is not None:
        anchor.addprevious(el)
    else:
        settings_root.append(el)
    return True


# ---------------------------------------------------------------------------
# sectPr 组装
# ---------------------------------------------------------------------------
def build_sectPr(template, margins, footer_default=None, footer_even=None,
                 header_default=None, header_even=None, pgnum=None):
    """
    基于 template（现有 sectPr 元素）复制一份，替换 pgMar 并挂上页眉页脚引用与 pgNumType。
    返回新的 sectPr 元素。

    注意：打开 evenAndOddHeaders 后，偶数页只认 even 类型的页眉页脚引用；
    因此页脚必须 default 与 even 同时挂（同一部件），否则偶数页会没有页码。
    """
    sec = etree.fromstring(etree.tostring(template))
    # 清掉旧的引用与页码定义
    for tag in ("headerReference", "footerReference", "pgNumType"):
        for ch in sec.findall(w(tag)):
            sec.remove(ch)
    # 页眉页脚引用须排在 sectPr 最前（CT_SectPr 顺序）
    ordered = []
    if header_even:
        e = etree.Element(w("headerReference"))
        e.set(w("type"), "even")
        e.set(r("id"), header_even)
        ordered.append(e)
    if header_default:
        e = etree.Element(w("headerReference"))
        e.set(w("type"), "default")
        e.set(r("id"), header_default)
        ordered.append(e)
    if footer_even:
        e = etree.Element(w("footerReference"))
        e.set(w("type"), "even")
        e.set(r("id"), footer_even)
        ordered.append(e)
    if footer_default:
        e = etree.Element(w("footerReference"))
        e.set(w("type"), "default")
        e.set(r("id"), footer_default)
        ordered.append(e)
    for i, el in enumerate(ordered):
        sec.insert(i, el)
    # 页面边距
    mar = sec.find(w("pgMar"))
    if mar is not None:
        for k, v in margins.items():
            mar.set(w(k), str(v))
    # 页码：pgNumType 在 CT_SectPr 中排于 pgMar 之后、cols/docGrid 之前
    if pgnum:
        pn = etree.Element(w("pgNumType"))
        for k, v in pgnum.items():
            pn.set(w(k), v)
        after_pgnum = ("cols", "formProt", "vAlign", "noEndnote", "titlePg",
                       "textDirection", "bidi", "rtlGutter", "docGrid",
                       "printerSettings", "sectPrChange")
        anchor = None
        for tag in after_pgnum:
            anchor = sec.find(w(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(pn)
        else:
            sec.append(pn)
    return sec


def set_para_sectpr(p, sectpr):
    """把 sectPr 写进段落 pPr（须为最后子元素）。"""
    ppr = p.find(w("pPr"))
    if ppr is None:
        ppr = etree.Element(w("pPr"))
        p.insert(0, ppr)
    old = ppr.find(w("sectPr"))
    if old is not None:
        ppr.remove(old)
    change = ppr.find(w("pPrChange"))
    if change is not None:
        change.addprevious(sectpr)
    else:
        ppr.append(sectpr)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def run(in_path, out_path, title, log, force=False, school_header=None):
    if os.path.abspath(in_path) == os.path.abspath(out_path):
        raise SystemExit("输出路径不能与原路径相同（请用 -o 指定新文件）")
    parts = kit.load_parts(in_path)
    # 幂等保护：已含页眉页脚部件的文档，重复执行只会不断新增部件
    existing_hf = [n for n in parts if re.search(r"word/(header|footer)\d+\.xml$", n)]
    if existing_hf and not force:
        raise SystemExit(
            "该文档已存在页眉页脚部件（%s），重复执行会产生冗余部件。\n"
            "确认要重建请加 --force。" % ", ".join(sorted(existing_hf)))
    doc_root = kit.parse_xml(parts["word/document.xml"])
    styles = kit.StyleTable(kit.parse_xml(parts["word/styles.xml"]))
    body = doc_root.find(w("body"))

    hdr_sid = styles.find_by_name("header") or styles.find_by_name("页眉") or "a3"
    ftr_sid = styles.find_by_name("footer") or styles.find_by_name("页脚") or "a5"

    blocks = list(kit.iter_blocks(body))
    front_no, body_no = find_boundaries(blocks, styles)
    log.append("分节定位：前置起始块 = %s，正文起始块 = %s" % (front_no, body_no))
    if front_no is None:
        raise SystemExit("未找到「摘要/Abstract/目录/声明」等前置标题，无法分节；请人工确认文档结构。")
    if body_no is None:
        raise SystemExit("未找到章标题（heading 1），无法定位正文起始；请人工确认。")

    # 前置 / 正文 的结束段落（分节符落点）
    def last_p_before(no):
        last = None
        for kind, el, n in blocks:
            if kind == "p" and n < no:
                last = el
        return last

    cover_end = last_p_before(front_no)
    front_end = last_p_before(body_no)
    if cover_end is None or front_end is None:
        raise SystemExit("分节落点定位失败。")

    chap = first_chapter_text(blocks, styles, body_no)

    # 第 1 节是否真的是「封面」——用与 fmt_check 相同的判据，避免两边口径不一致。
    # 若不是封面（例如缺少封面字样），就沿用正文页边距，不做无谓的版式改动。
    cover_text = kit.range_text(blocks, None, front_no - 1)
    is_cover = kit.section_kind(cover_text) == "cover"
    sec1_margins = spec.MARGIN_COVER if is_cover else spec.MARGIN_BODY
    log.append("第 1 节类型：%s（页边距按%s）" % (kit.section_kind(cover_text),
                                              "封面" if is_cover else "正文"))

    # --- 页眉页脚部件 ---
    used = {n.split("/")[-1] for n in parts if re.search(r"word/(header|footer)\d+\.xml$", n)}
    rels_root = kit.parse_xml(parts["word/_rels/document.xml.rels"])
    ct_root = kit.parse_xml(parts["[Content_Types].xml"])

    ftr_page_rid = add_part(parts, rels_root, ct_root, "footer",
                            make_footer_page(ftr_sid, "1"), used)
    if school_header:
        # 全校统一页眉（如「XX大学本科生毕业论文」）：奇偶页通用同一部件
        hdr_odd_rid = add_part(parts, rels_root, ct_root, "header",
                               make_header_text(hdr_sid, school_header), used)
        hdr_even_rid = hdr_odd_rid
        log.append("页眉：使用统一校名页眉「%s」（奇偶页相同）" % school_header)
    else:
        hdr_odd_rid = add_part(parts, rels_root, ct_root, "header",
                               make_header_styleref(hdr_sid, chap), used)
        hdr_even_rid = add_part(parts, rels_root, ct_root, "header",
                                make_header_text(hdr_sid, title), used)
        log.append("页眉：奇数页 STYLEREF 章名 / 偶数页论文题目")

    body_sect = body.find(w("sectPr"))
    if body_sect is None:
        body_sect = etree.SubElement(body, w("sectPr"))

    # --- 第 1 节：封面（无页眉页脚）---
    sec1 = build_sectPr(body_sect, sec1_margins)
    set_para_sectpr(cover_end, sec1)
    # --- 第 2 节：前置（罗马页码 + 页码页脚）---
    sec2 = build_sectPr(body_sect, spec.MARGIN_BODY,
                        footer_default=ftr_page_rid, footer_even=ftr_page_rid,
                        pgnum={"fmt": spec.PGNUM_FRONT, "start": "1"})
    set_para_sectpr(front_end, sec2)
    # --- 第 3 节：正文（奇偶页页眉 + 阿拉伯页码重启）---
    sec3 = build_sectPr(body_sect, spec.MARGIN_BODY,
                        footer_default=ftr_page_rid, footer_even=ftr_page_rid,
                        header_default=hdr_odd_rid, header_even=hdr_even_rid,
                        pgnum={"fmt": spec.PGNUM_BODY, "start": "1"})
    body.replace(body_sect, sec3)

    # --- settings ---
    settings_root = kit.parse_xml(parts["word/settings.xml"])
    if insert_even_odd_headers(settings_root):
        log.append("settings：已打开 evenAndOddHeaders（奇偶页页眉不同）")
    if ensure_update_fields(settings_root):
        log.append("settings：已打开 updateFields（Word 打开时自动刷新域）")

    # --- 写回 ---
    out_parts = dict(parts)
    out_parts["word/document.xml"] = kit.serialize_xml(doc_root)
    out_parts["word/settings.xml"] = kit.serialize_xml(settings_root)
    out_parts["word/_rels/document.xml.rels"] = kit.serialize_xml(rels_root)
    out_parts["[Content_Types].xml"] = kit.serialize_xml(ct_root)

    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in out_parts.items():
            z.writestr(n, data)
    log.append("已写出：%s" % out_path)
    log.append("页眉页脚部件：1 个页脚 + 2 个页眉（奇数页 STYLEREF 章名 / 偶数页论文题目）")
    log.append("页码：封面无 → 前置大写罗马(I 起) → 正文阿拉伯(1 起)")
    return log


def main():
    ap = argparse.ArgumentParser(description="为已规范化的论文建立分节、页码与页眉")
    ap.add_argument("docx", help="输入 docx（应先经 fmt_apply 规范化）")
    ap.add_argument("-o", "--out", required=True, help="输出 docx 路径")
    ap.add_argument("--title", default="", help="论文题目（偶数页页眉用）")
    ap.add_argument("--report", default=None, help="把变更日志写到该文件")
    ap.add_argument("--force", action="store_true", help="文档已有页眉页脚时仍强制重建")
    ap.add_argument("--school-header", default=None,
                    help="改用统一的固定页眉文本（如「XX大学本科生毕业论文」），奇偶页相同")
    args = ap.parse_args()

    log = []
    run(args.docx, args.out, args.title, log, force=args.force,
        school_header=args.school_header)
    text = "\n".join(log)
    print(text)
    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            f.write(text + "\n")


if __name__ == "__main__":
    main()
