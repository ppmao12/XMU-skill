# -*- coding: utf-8 -*-
"""
封面品牌替换 —— 把封面上的校徽 / 校名图片换成指定学校的标志。

适用场景：论文模板是 A 校的（封面左上角是 A 校校徽 + 校名书法字），
要改造成 B 校。版面（图片位置、下方「本科毕业论文」字样、信息表）保持不变，
只换图形本身。

实现要点：
  * 封面图区 = 文档开头、第一段有正文的段落之前，所有含图片的段落。
  * 只保留第一张图片（通常占位最宽，即校名那一条），把它的位图换成新标志，
    并按新标志的宽高比重算 width/height，避免拉伸变形；其余封面图整段删除。
  * 图片尺寸用 pt 写在 VML 的 style 里，Word 按 pt 缩放，所以无需重采样位图。
  * 保持原图扩展名（png→png、jpeg→jpeg），这样 [Content_Types].xml 无需改动。

用法：
    python fmt_cover.py <封面含旧校徽的.docx> -o <输出.docx> --logo <新标志.png>
                       [--width-pt 300] [--top-pt 28] [--report 日志.txt]
"""
import argparse
import io
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
import docx_kit as kit
import spec_rules as spec

VML_NS = "urn:schemas-microsoft-com:vml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def q(ns, tag):
    return "{%s}%s" % (ns, tag)


# ---------------------------------------------------------------------------
# 定位封面图区
# ---------------------------------------------------------------------------
def cover_image_paragraphs(body):
    """
    返回封面图区里的段落（按出现顺序）。

    规则：从文档开头逐段扫描，遇到「有文字且无图」的段落即认为封面图区结束
    （封面图区总是位于「本科毕业论文 / 硕士学位论文」等字样之前）。
    """
    out = []
    for el in body:
        tag = etree.QName(el).localname
        if tag == "tbl":
            continue
        if tag != "p":
            continue
        has_img = any(etree.QName(n).localname in ("imagedata", "blip")
                      for n in el.iter())
        if not has_img and kit.para_text(el).strip():
            break
        if has_img:
            out.append(el)
    return out


def collect_images(paras):
    """返回 [(para, run, holder, node, rid)]，holder 为承载尺寸的 v:shape 或 wp:inline。"""
    found = []
    for p in paras:
        for node in p.iter():
            ln = etree.QName(node).localname
            if ln == "imagedata":
                rid = node.get(q(R_NS, "id")) or node.get(q(R_NS, "embed"))
                if rid:
                    holder = node.getparent()
                    while holder is not None and etree.QName(holder).localname != "shape":
                        holder = holder.getparent()
                    found.append((p, node.getparent(), holder, node, rid))
            elif ln == "blip":
                rid = node.get(q(R_NS, "embed"))
                if rid:
                    holder = node.getparent()
                    while holder is not None and etree.QName(holder).localname not in ("inline", "anchor"):
                        holder = holder.getparent()
                    found.append((p, node.getparent(), holder, node, rid))
    return found


# ---------------------------------------------------------------------------
# 位图读写
# ---------------------------------------------------------------------------
def load_logo(path):
    from PIL import Image
    im = Image.open(path)
    return im


def encode_like(im, target_part):
    """把新标志编码成与目标部件相同的格式（保持 [Content_Types] 不变）。"""
    buf = io.BytesIO()
    ext = os.path.splitext(target_part)[1].lower()
    if ext in (".jpg", ".jpeg"):
        if im.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            rgba = im.convert("RGBA")
            bg.paste(rgba, mask=rgba.split()[-1])
            im = bg
        im.save(buf, "JPEG", quality=95)
    elif ext == ".png":
        if im.mode not in ("RGBA", "RGB", "L", "P"):
            im = im.convert("RGBA")
        im.save(buf, "PNG")
    else:
        im.save(buf, "PNG")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# VML style 串处理
# ---------------------------------------------------------------------------
def parse_style(s):
    items = []
    for part in (s or "").split(";"):
        if not part.strip():
            continue
        if ":" in part:
            k, v = part.split(":", 1)
            items.append((k.strip(), v.strip()))
        else:
            items.append((part.strip(), ""))
    return items


def dump_style(items):
    return ";".join("%s:%s" % (k, v) if v != "" else k for k, v in items)


def set_style(items, key, value):
    for i, (k, v) in enumerate(items):
        if k == key:
            items[i] = (key, value)
            return
    items.append((key, value))


def get_style(items, key, default=None):
    for k, v in items:
        if k == key:
            return v
    return default


# ---------------------------------------------------------------------------
# 版心宽度
# ---------------------------------------------------------------------------
def usable_width_pt(parts):
    """取第一个 sectPr 的版心宽度（pt）。"""
    doc = kit.parse_xml(parts["word/document.xml"])
    body = doc.find(kit.w("body"))
    sect = body.find(kit.w("sectPr"))
    if sect is None:
        for el in body.iter():
            if etree.QName(el).localname == "sectPr":
                sect = el
                break
    if sect is None:
        return 0.0, 0.0, 0.0
    sz = sect.find(kit.w("pgSz"))
    mar = sect.find(kit.w("pgMar"))
    w = int(sz.get(kit.w("w"))) if sz is not None else 11906
    left = int(mar.get(kit.w("left"))) if mar is not None else 0
    right = int(mar.get(kit.w("right"))) if mar is not None else 0
    return (w - left - right) / 20.0, left / 20.0, right / 20.0


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def run(in_path, out_path, logo_path, log, width_pt=None, top_pt=None,
        align="center", left_pt=None):
    if os.path.abspath(in_path) == os.path.abspath(out_path):
        raise SystemExit("输出路径不能与原路径相同（请用 -o 指定新文件）")

    parts = kit.load_parts(in_path)
    doc_root = kit.parse_xml(parts["word/document.xml"])
    body = doc_root.find(kit.w("body"))
    rels = kit.parse_rels(parts, "word/document.xml")

    paras = cover_image_paragraphs(body)
    found = collect_images(paras)
    if not found:
        raise SystemExit("未在封面区找到图片；请确认该文档是「封面带校徽/校名图」的模板")

    log.append("封面图区：%d 个段落 / %d 张图片" % (len(paras), len(found)))

    im = load_logo(logo_path)
    aspect = im.width / float(im.height)

    keep = found[0]
    drop = found[1:]
    log.append("保留首图并替换为新标志；其余 %d 张移除" % len(drop))

    # ---- 替换位图 ----
    # 注意：docx_kit.parse_rels 返回的已经是「补全目录后」的部件名（如 word/media/image1.png）
    part = rels[keep[4]].replace("\\", "/").lstrip("/")
    if part not in parts:
        raise SystemExit("找不到图片部件：%s" % part)
    old_size = len(parts[part])
    parts[part] = encode_like(im, part)
    log.append("  %s ← %s（%d → %d 字节）"
               % (part, os.path.basename(logo_path), old_size, len(parts[part])))

    # ---- 尺寸与位置 ----
    usable, ml, mr = usable_width_pt(parts)
    sec = body.find(kit.w("sectPr"))
    if sec is not None:
        sz = sec.find(kit.w("pgSz"))
        mar = sec.find(kit.w("pgMar"))
        if sz is not None and mar is not None:
            usable = (int(sz.get(kit.w("w"))) - int(mar.get(kit.w("left")))
                      - int(mar.get(kit.w("right")))) / 20.0

    holder = keep[2]
    if holder is not None and etree.QName(holder).localname == "shape":
        items = parse_style(holder.get("style"))
        cur_w = get_style(items, "width", "284pt")
        try:
            cur_w_pt = float(re.sub(r"[^0-9.]", "", cur_w) or 284)
        except ValueError:
            cur_w_pt = 284.0
        new_w = float(width_pt) if width_pt else min(cur_w_pt, usable)
        new_h = new_w / aspect
        set_style(items, "width", "%.2fpt" % new_w)
        set_style(items, "height", "%.2fpt" % new_h)
        if top_pt is not None:
            set_style(items, "margin-top", "%.2fpt" % float(top_pt))
        if align == "center":
            set_style(items, "margin-left", "%.2fpt" % max(0.0, (usable - new_w) / 2.0))
        elif left_pt is not None:
            set_style(items, "margin-left", "%.2fpt" % float(left_pt))
        holder.set("style", dump_style(items))
        log.append("  尺寸 → %.1fpt × %.1fpt（版心 %.1fpt，居中）" % (new_w, new_h, usable))
    elif holder is not None:
        # DrawingML：改 wp:extent
        ext = holder.find(q(WP_NS, "extent"))
        if ext is not None:
            cur_w = int(ext.get("cx")) / 12700.0  # EMU -> pt
            new_w = float(width_pt) if width_pt else min(cur_w, usable)
            new_h = new_w / aspect
            ext.set("cx", str(int(new_w * 12700)))
            ext.set("cy", str(int(new_h * 12700)))
            log.append("  尺寸（DrawingML）→ %.1fpt × %.1fpt" % (new_w, new_h))

    # ---- 删除多余封面图 ----
    for p, run_el, holder2, node, rid in drop:
        parent = holder2 if holder2 is not None else node
        run_parent = parent
        while run_parent is not None and etree.QName(run_parent).localname != "r":
            run_parent = run_parent.getparent()
        target = run_parent if run_parent is not None else parent
        if target.getparent() is not None:
            target.getparent().remove(target)
        # 清掉空 run
        for r in list(p.findall(kit.w("r"))):
            rest = [c for c in r if etree.QName(c).localname != "rPr"]
            if not rest:
                p.remove(r)
        log.append("  移除多余封面图（rId=%s）" % rid)

    # ---- 先摘掉这些图片的关系，再清理不再被引用的媒体部件 ----
    rels_root = kit.parse_xml(parts["word/_rels/document.xml.rels"])
    dropped_targets = set()
    for p2, run_el, holder2, node, rid in drop:
        tgt = rels.get(rid)
        if tgt:
            dropped_targets.add(tgt)
        for rel in list(rels_root):
            if rel.get("Id") == rid:
                rels_root.remove(rel)
    # 用「摘掉关系之后」的关系表判断是否还有引用，避免把仍在用的图删掉
    rels_text = kit.serialize_xml(rels_root).decode("utf-8", "ignore")
    removed = 0
    for tgt in dropped_targets:
        media = tgt.replace("\\", "/").lstrip("/")
        base = os.path.basename(media)
        if media in parts and base not in rels_text:
            del parts[media]
            removed += 1
    log.append("  清理孤立媒体部件：%d 个" % removed)

    # ---- 写回 ----
    out_parts = dict(parts)
    out_parts["word/document.xml"] = kit.serialize_xml(doc_root)
    out_parts["word/_rels/document.xml.rels"] = kit.serialize_xml(rels_root)

    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in out_parts.items():
            z.writestr(n, data)
    log.append("已写出：%s" % out_path)
    return log


def main():
    ap = argparse.ArgumentParser(description="把封面上的校徽/校名图换成指定学校的标志")
    ap.add_argument("docx", help="输入 docx")
    ap.add_argument("-o", "--out", required=True, help="输出 docx 路径")
    ap.add_argument("--logo", required=True, help="新标志图片（建议官方 logo，PNG）")
    ap.add_argument("--width-pt", type=float, default=None,
                    help="新标志宽度（pt）；缺省沿用原图宽度")
    ap.add_argument("--top-pt", type=float, default=None, help="距页顶（pt）")
    ap.add_argument("--left-pt", type=float, default=None, help="左偏移（pt，align=left 时生效）")
    ap.add_argument("--align", default="center", choices=["center", "left"],
                    help="水平对齐：center（默认）或 left")
    ap.add_argument("--report", default=None, help="把变更日志写到该文件")
    args = ap.parse_args()

    log = []
    run(args.docx, args.out, args.logo, log, width_pt=args.width_pt,
        top_pt=args.top_pt, align=args.align, left_pt=args.left_pt)
    text = "\n".join(log)
    print(text)
    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            f.write(text + "\n")


if __name__ == "__main__":
    main()
