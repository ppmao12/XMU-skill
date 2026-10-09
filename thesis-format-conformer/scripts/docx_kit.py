# -*- coding: utf-8 -*-
"""
OOXML 处理工具库（thesis-format-conformer 专用）

提供：
  - 读取 .docx（zip 容器）与部件
  - 解析 styles.xml，解析 basedOn 样式继承链，计算"有效格式"
  - 识别段落语义角色（章/节/小节标题、图表题、参考文献……）
  - 解析分节（sectPr）：页面尺寸、页边距、页码格式、页眉页脚引用

设计立场：把 .docx 当作 OOXML（zip 内 XML 部件）处理，不做纯文本替换。
所有取样仅做只读解析；写操作由 fmt_apply.py 负责。
"""
import os
import re
import zipfile

from lxml import etree

# ---------------------------------------------------------------------------
# 命名空间
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"

NSMAP = {
    "w": W_NS,
    "r": R_NS,
    "w14": W14_NS,
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def w(tag):
    return "{%s}%s" % (W_NS, tag)


def rel_id(el):
    return el.get("{%s}id" % R_NS)


# ---------------------------------------------------------------------------
# 载入 / 保存
# ---------------------------------------------------------------------------
def load_parts(docx_path):
    """读取 docx 全部部件为 {部件名: bytes}。"""
    parts = {}
    with zipfile.ZipFile(docx_path, "r") as z:
        for name in z.namelist():
            parts[name] = z.read(name)
    return parts


def parse_xml(data):
    """解析 XML 字节流，返回 lxml 元素（保留注释与处理指令）。"""
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    return etree.fromstring(data, parser=parser)


def serialize_xml(root, declaration=True):
    return etree.tostring(root, xml_declaration=declaration, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
# 格式属性抽取
# ---------------------------------------------------------------------------
# 统一属性字典（"有效格式"）：
#   jc / firstLine / firstLineChars / left / leftChars / hanging
#   line / lineRule / before / after
#   ascii / eastAsia / hAnsi / sz / bold / italic / outlineLvl / pageBreakBefore


def _int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def props_from_ppr(ppr):
    """从 w:pPr 抽取段落级格式。"""
    out = {}
    if ppr is None:
        return out
    jc = ppr.find(w("jc"))
    if jc is not None:
        out["jc"] = jc.get(w("val")) or "left"
    ind = ppr.find(w("ind"))
    if ind is not None:
        for a in ("firstLine", "firstLineChars", "left", "leftChars", "hanging", "right", "rightChars"):
            v = _int(ind.get(w(a)))
            if v is not None:
                out[a] = v
    sp = ppr.find(w("spacing"))
    if sp is not None:
        for a in ("line", "before", "after", "beforeLines", "afterLines"):
            v = _int(sp.get(w(a)))
            if v is not None:
                out[a] = v
        lr = sp.get(w("lineRule"))
        if lr:
            out["lineRule"] = lr
    ol = ppr.find(w("outlineLvl"))
    if ol is not None:
        v = _int(ol.get(w("val")))
        if v is not None:
            out["outlineLvl"] = v
    if ppr.find(w("pageBreakBefore")) is not None:
        out["pageBreakBefore"] = True
    if ppr.find(w("keepNext")) is not None:
        out["keepNext"] = True
    # 段落标记自身的 rPr（影响段落行高与空段落显示）
    rpr = ppr.find(w("rPr"))
    if rpr is not None:
        out.update(props_from_rpr(rpr))
    return out


def props_from_rpr(rpr):
    """从 w:rPr 抽取字符级格式。字号统一换算为 pt。"""
    out = {}
    if rpr is None:
        return out
    rf = rpr.find(w("rFonts"))
    if rf is not None:
        for a in ("ascii", "eastAsia", "hAnsi", "cs"):
            v = rf.get(w(a))
            if v:
                out[a] = v
    sz = rpr.find(w("sz"))
    if sz is not None:
        v = _int(sz.get(w("val")))
        if v is not None:
            out["sz"] = v / 2.0
    b = rpr.find(w("b"))
    if b is not None:
        out["bold"] = b.get(w("val")) not in ("0", "false", "off")
    i = rpr.find(w("i"))
    if i is not None:
        out["italic"] = i.get(w("val")) not in ("0", "false", "off")
    return out


def merge(base, override):
    """override 覆盖 base（None / 缺失不覆盖）。"""
    out = dict(base)
    for k, v in (override or {}).items():
        if v is not None:
            out[k] = v
    return out


# ---------------------------------------------------------------------------
# 样式表
# ---------------------------------------------------------------------------
class StyleInfo(object):
    __slots__ = ("style_id", "name", "type", "based_on", "pPr", "rPr", "_props")

    def __init__(self, el):
        self.style_id = el.get(w("styleId"))
        nm = el.find(w("name"))
        self.name = (nm.get(w("val")) if nm is not None else "") or ""
        self.type = el.get(w("type")) or "paragraph"
        bo = el.find(w("basedOn"))
        self.based_on = bo.get(w("val")) if bo is not None else None
        self.pPr = el.find(w("pPr"))
        self.rPr = el.find(w("rPr"))
        self._props = merge(props_from_ppr(self.pPr), props_from_rpr(self.rPr))
        # pPr 里内嵌的 rPr 也会被 props_from_ppr 合并，此处覆盖语义正确

    @property
    def props(self):
        return self._props


class StyleTable(object):
    def __init__(self, styles_root):
        self.root = styles_root
        self.styles = {}
        self.by_name = {}
        self.doc_defaults = {}
        dd = styles_root.find(w("docDefaults"))
        if dd is not None:
            rpd = dd.find(w("rPrDefault"))
            if rpd is not None and rpd.find(w("rPr")) is not None:
                self.doc_defaults.update(props_from_rpr(rpd.find(w("rPr"))))
            ppd = dd.find(w("pPrDefault"))
            if ppd is not None:
                self.doc_defaults.update(props_from_ppr(ppd.find(w("pPr"))))
        for el in styles_root.findall(w("style")):
            si = StyleInfo(el)
            self.styles[si.style_id] = si
            self.by_name[si.name.strip().lower()] = si.style_id

    def find_by_name(self, name):
        return self.by_name.get(name.strip().lower())

    def chain(self, style_id):
        """样式继承链：从最远祖先到自身。"""
        seq = []
        seen = set()
        cur = style_id
        while cur and cur in self.styles and cur not in seen:
            seen.add(cur)
            seq.append(cur)
            cur = self.styles[cur].based_on
        seq.reverse()
        return seq

    def effective_style_id(self, style_id):
        """把样式链合并为一份有效格式。"""
        props = dict(self.doc_defaults)
        for sid in self.chain(style_id):
            props = merge(props, self.styles[sid].props)
        return props


# ---------------------------------------------------------------------------
# 段落与块
# ---------------------------------------------------------------------------
def iter_inline(p):
    """
    按文档顺序产出段落的内容级行内元素 (kind, node)，kind ∈ {'t','tab','br'}。

    只取 run（w:r）的直接子元素 —— 这一点很关键：
    w:pPr/w:tabs 里的 w:tab 是**制表位定义**（只有 w:val/w:pos 属性），
    不是正文制表符。若像早期实现那样遍历 p.iter() 抓所有 w:tab，
    带制表位的段落会凭空多出一个前导 "\\t"，进而让所有按字符数
    定位的偏移量（如剥离手写编号）整体错位。
    """
    for r in p.iter(w("r")):
        for node in r:
            tag = node.tag
            if tag == w("t"):
                yield "t", node
            elif tag == w("tab"):
                yield "tab", node
            elif tag == w("br"):
                yield "br", node


def para_text(p, maxlen=None):
    parts = []
    for kind, node in iter_inline(p):
        if kind == "t":
            parts.append(node.text or "")
        elif kind == "tab":
            parts.append("\t")
        else:
            parts.append("\n")
    s = "".join(parts)
    if maxlen:
        s = s[:maxlen]
    return s


def strip_prefix_chars(p, n):
    """
    从段落开头按「字符数」删除 n 个字符，跨 run 正确推进。

    用途：剥离标题文本里的手写编号（如 "1.1 选题背景" → "选题背景"），
    让编号改由样式/段落级自动编号生成。返回实际删除的字符数。

    计数口径必须与 para_text() 完全一致：w:t 按文本长度计，
    w:tab / w:br 各计 1 个字符。两者共用 iter_inline()，
    因此不会出现「制表位定义被误计」导致的整体错位（实测踩过）。
    """
    if n <= 0:
        return 0
    removed = 0
    for kind, node in iter_inline(p):
        if removed >= n:
            break
        if kind == "t":
            s = node.text or ""
            take = min(len(s), n - removed)
            if not take:
                continue
            rest = s[take:]
            node.text = rest
            removed += take
            # 保留首尾空白，否则 Word 会把它裁掉
            if rest != rest.strip():
                node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        else:
            removed += 1
            node.getparent().remove(node)
    return removed


def run_props_mode(p):
    """统计段落内 run 的字体/字号众数（返回出现最多的属性组合）。"""
    from collections import Counter
    cnt = Counter()
    for r in p.findall(w("r")):
        pr = props_from_rpr(r.find(w("rPr")))
        if not pr:
            continue
        key = (pr.get("eastAsia"), pr.get("ascii"), pr.get("sz"), pr.get("bold"))
        cnt[key] += len(para_text(r)) or 1
    if not cnt:
        return {}
    key = cnt.most_common(1)[0][0]
    out = {}
    if key[0]:
        out["eastAsia"] = key[0]
    if key[1]:
        out["ascii"] = key[1]
    if key[2] is not None:
        out["sz"] = key[2]
    if key[3] is not None:
        out["bold"] = key[3]
    return out


def para_effective(p, styles):
    """段落有效格式 = docDefaults → 样式链 → 段落直接格式。"""
    ppr = p.find(w("pPr"))
    style_id = None
    if ppr is not None:
        ps = ppr.find(w("pStyle"))
        if ps is not None:
            style_id = ps.get(w("val"))
    props = styles.effective_style_id(style_id) if style_id else dict(styles.doc_defaults)
    props = merge(props, props_from_ppr(ppr))
    return props, style_id


def style_name_of(styles, style_id):
    si = styles.styles.get(style_id) if style_id else None
    return si.name if si else None


# ---------------------------------------------------------------------------
# 角色识别
# ---------------------------------------------------------------------------
RE_CHAPTER = re.compile(r"^第\s*[一二三四五六七八九十百零〇0-9]+\s*章(\s|$|[\u4e00-\u9fff])")
RE_SEC = re.compile(r"^(\d+)\.(\d+)(?!\.)\s*\S")
RE_SUBSEC = re.compile(r"^(\d+)\.(\d+)\.(\d+)(?!\.)\s*\S")
RE_SUBSUB = re.compile(r"^(\d+)\.(\d+)\.(\d+)\.(\d+)\s*\S")
RE_FIG = re.compile(r"^图\s*\d+([\.\-–]\d+)?\s*\S")
RE_TAB = re.compile(r"^表\s*\d+([\.\-–]\d+)?\s*\S")
RE_NOTE = re.compile(r"^注\s*[:：]")
# 提纲式纯文本编号标题：「1 金融监管的必要性」「2 面临的风险」
# 限定 1-2 位数字 + 空白，避免把「2023年3月，…」这类正文误判
RE_OUTLINE = re.compile(r"^\s*\d{1,2}[ \u3000]+\S")
RE_KEYWORDS = re.compile(r"^(关键词|关键字|Keywords?)\s*[:：]", re.I)
RE_REF_MARK = re.compile(r"\[(J|M|D|C|R|N|P|S|A|EB/OL|Z)\]", re.I)

MAJOR_TITLES = {
    "摘要", "摘 要", "摘  要", "中文摘要",
    "abstract", "英文摘要",
    "目录", "目 录", "目  录", "contents",
    "参考文献", "references",
    "致谢", "acknowledgement", "acknowledgments", "谢辞",
    "附录", "appendix",
}

ROLE_STYLE_NAMES = {
    "表标题": "table_caption",
    "图标题": "figure_caption",
    "图表说明": "table_note",
    "表格内文字": "table_body",
    "文献列表": "reference_item",
    "论文标题": "cover_subject",
    "论文题目": "cover_subject",
    "内容摘要": "major_title",
}


def detect_role(text, props, styles, style_id=None, in_reference_block=False, sect_kind="body"):
    """
    识别段落语义角色。返回 ROLE 字符串（见 spec_rules.ROLE_TO_STYLE）。

    判据优先级：目录条目 > 节类型（封面/声明） > 显式样式名 > 文本模式 > 大纲级别 > 兜底正文。
    """
    stripped = text.strip()
    compact = stripped.replace(" ", "").replace("\u3000", "")
    sname = style_name_of(styles, style_id) or ""
    sname_l = sname.strip().lower()
    lower = compact.lower()

    # 0) 目录条目 —— 自动生成的域，不属于正文规范管辖
    if sname_l.startswith("toc"):
        return "toc_entry"

    # 1) 空段落 —— 不参与格式判定（含仅用于分页的空段落）
    if not stripped:
        return "blank"

    # 2) 封面 / 声明页 —— 自由排版区，不套正文规范
    if sect_kind == "cover":
        return "cover"
    if sect_kind == "declaration":
        return "declaration"

    # 3) 显式命名样式（最高置信度）
    if sname in ROLE_STYLE_NAMES:
        return ROLE_STYLE_NAMES[sname]
    if sname_l in ("我的标题",):
        return "major_title"
    if sname_l.startswith("heading 1") or sname_l == "标题 1":
        return "chapter_title"
    if sname_l.startswith("heading 2") or sname_l == "标题 2":
        return "section_title"
    if sname_l.startswith("heading 3") or sname_l == "标题 3":
        return "subsection_title"
    if sname_l.startswith("heading 4") or sname_l == "标题 4":
        return "subsubsection_title"

    # 2) 篇级大标题
    if not stripped:
        return "blank"
    if lower in MAJOR_TITLES:
        return "major_title"
    # "摘 要" 之类带空格的
    if compact in {"摘要", "目录", "参考文献", "致谢", "附录"} and len(stripped) <= 6:
        return "major_title"

    # 3) 章标题
    if RE_CHAPTER.match(stripped) and len(stripped) <= 40:
        return "chapter_title"

    # 4) 图题 / 表题（长度限制，避免误伤正文引用）
    if len(stripped) <= 80 and RE_FIG.match(stripped):
        return "figure_caption"
    if len(stripped) <= 80 and RE_TAB.match(stripped):
        return "table_caption"

    # 5) 图表注
    if RE_NOTE.match(stripped) and len(stripped) <= 400:
        return "table_note"

    # 6) 编号型标题（文本独立成行、较短、无句末句号）
    if len(stripped) <= 60 and not stripped.endswith(("。", "；", ".", ";")):
        if RE_SUBSUB.match(stripped):
            return "subsubsection_title"
        if RE_SUBSEC.match(stripped):
            return "subsection_title"
        if RE_SEC.match(stripped):
            return "section_title"
        # 提纲式纯文本编号标题（初稿常见：正文里手写成 "1 面临的风险"）
        # 严格限制长度，且已由上方的句末标点条件挡住带标点的正文句
        if RE_OUTLINE.match(stripped) and len(compact) <= 40:
            return "section_title"

    # 7) 关键词行
    if RE_KEYWORDS.match(stripped):
        return "keywords"

    # 8) 参考文献条目
    if in_reference_block and RE_REF_MARK.search(stripped):
        return "reference_item"
    if in_reference_block and len(stripped) > 20:
        return "reference_item"

    # 9) 依据大纲级别兜底
    ol = props.get("outlineLvl")
    if ol is not None:
        return {
            0: "chapter_title",
            1: "section_title",
            2: "subsection_title",
            3: "subsubsection_title",
        }.get(ol, "body")

    return "body"


# ---------------------------------------------------------------------------
# 分节解析
# ---------------------------------------------------------------------------
class Section(object):
    def __init__(self, sectpr):
        self.el = sectpr
        self.pg_w = self.pg_h = None
        self.margins = {}
        self.headers = {}
        self.footers = {}
        self.pgnum_fmt = None
        self.pgnum_start = None
        self.doc_grid_pitch = None
        self.title_pg = False
        pg = sectpr.find(w("pgSz"))
        if pg is not None:
            self.pg_w = _int(pg.get(w("w")))
            self.pg_h = _int(pg.get(w("h")))
        mar = sectpr.find(w("pgMar"))
        if mar is not None:
            for a in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
                v = _int(mar.get(w(a)))
                if v is not None:
                    self.margins[a] = v
        for ref in sectpr.findall(w("headerReference")):
            rid = ref.get(w("type"))
            self.headers[rid] = rel_id(ref)
        for ref in sectpr.findall(w("footerReference")):
            rid = ref.get(w("type"))
            self.footers[rid] = rel_id(ref)
        pn = sectpr.find(w("pgNumType"))
        if pn is not None:
            self.pgnum_fmt = pn.get(w("fmt"))
            self.pgnum_start = _int(pn.get(w("start")))
        dg = sectpr.find(w("docGrid"))
        if dg is not None:
            self.doc_grid_pitch = _int(dg.get(w("linePitch")))
        if sectpr.find(w("titlePg")) is not None:
            self.title_pg = True


def iter_blocks(body):
    """
    按文档顺序产出 body 的直接子块。
    产出 (kind, element, para_no)：
      kind: 'p' | 'tbl' | 'sectPr'
    para_no 为连续计数（段落与表格都计入），便于报告定位。
    """
    n = 0
    for child in body:
        tag = child.tag
        if tag == w("p"):
            n += 1
            yield "p", child, n
        elif tag == w("tbl"):
            n += 1
            yield "tbl", child, n
        elif tag == w("sectPr"):
            yield "sectPr", child, n


def collect_sections(body):
    """
    返回 [(Section, start_para, end_para)]，按出现顺序。

    注意：Word 的分节符有两种落点 ——
      (a) 段落的 pPr 内嵌 w:sectPr（"下一页分节符"最常用）；
      (b) body 直接子元素 w:sectPr（文档最后一节）。
    两者都要识别，否则会把多节文档误判为单节。
    """
    secs = []
    start = 1
    n = 0
    for child in body:
        tag = child.tag
        if tag == w("p"):
            n += 1
            ppr = child.find(w("pPr"))
            if ppr is not None:
                sp = ppr.find(w("sectPr"))
                if sp is not None:
                    secs.append((Section(sp), start, n))
                    start = n + 1
        elif tag == w("tbl"):
            n += 1
        elif tag == w("sectPr"):
            secs.append((Section(child), start, n))
            start = n + 1
    return secs


# ---------------------------------------------------------------------------
# 节类型判定（封面 / 声明页 / 前置部分 / 主体）
# ---------------------------------------------------------------------------
RE_COVER_HINT = re.compile(
    r"学校编码|学\s*号|硕\s*士\s*学\s*位\s*论\s*文|指导教师|论文提交日期|论文答辩日期|"
    r"学位授予日期|培养单位|专业名称|研究方向|作者姓名")
RE_DECL_HINT = re.compile(r"原创性声明|著作权使用声明|学术不端|保密委员会|学位论文著作权")
# 章标题（第X章 / 第X篇）—— 出现它即说明该节已进入论文主体
RE_BODY_START = re.compile(r"第\s*[一二三四五六七八九十百零〇0-9]+\s*[章篇]")
# 「前置部分」起始标记：摘要 / Abstract / 目录 / 声明 / 参考文献 / 致谢 / 附录。
# 用于从文档开头裁出封面区（见 leading_front_end）。
RE_FRONT_START = re.compile(
    r"^\s*(摘\s*要|abstract|目\s*录|contents|原创性声明|著作权使用声明|"
    r"参考文献|references|致\s*谢|附\s*录)", re.I)
# 封面/声明字样只会出现在本节开头，故判型时只看开头这些字符
HEAD_LOOK = 400
# 一节文本长到这个量级就不可能是封面/声明页了
BODY_TEXT_MIN = 1200


def range_text(blocks, lo=None, hi=None):
    """
    拼接块号区间内**所有**块（含表格）的文本。

    用途：判定节的类型（封面 / 声明 / 前置 / 正文）。
    必须把**表格单元格**里的文字也算上 —— 封面的「学校编码 / 学号 / 作者姓名 / 指导教师」
    等信息通常放在封面信息表里；只看段落文本会把封面误判为正文节，
    进而给封面套上正文页边距、并让封面的段落被当成正文去套样式（实测踩过）。
    """
    out = []
    for kind, el, no in blocks:
        if lo is not None and no < lo:
            continue
        if hi is not None and no > hi:
            continue
        if kind == "p":
            out.append(para_text(el))
        elif kind == "tbl":
            for p in el.iter(w("p")):
                out.append(para_text(p))
    return "".join(out)


def _looks_like_body(text):
    """
    整节文本里是否已出现「论文主体」的迹象。

    判据只有两条，都刻意从严（宁可判成正文）：
      * 出现章标题（第X章 / 第X篇）；
      * 文本长度 ≥ BODY_TEXT_MIN 字符 —— 封面与声明页远达不到这个量级。

    为什么需要这道闸：源稿（尤其 .doc 转换来的、尚未分节的）常把封面与正文
    放在**同一节**里。此时整节文本既有「学号 / 指导教师」等封面字样，又有整篇
    正文。若不加闸，整节会被判成封面 → 所有段落一律跳过规范化（实测 109 个
    段落全部漏做），封面标题也会被顺带压成小四。
    """
    if RE_BODY_START.search(text):
        return True
    return len(text) >= BODY_TEXT_MIN


def section_kind(text_all):
    """
    依据整节文本判定节的类型：
      cover       —— 封面（自由排版，不套正文规范）
      declaration —— 原创性声明 / 著作权使用声明页
      front       —— 摘要、Abstract、目录、Contents
      body        —— 正文及其后

    两道约束：
      1) 封面/声明字样只会出现在本节开头，因此只看开头 HEAD_LOOK 个字符
         （避免正文里偶然出现的「研究方向」「学术不端」把整节带走）；
      2) 先用 _looks_like_body() 拦一道 —— 本节若已是论文主体，就一律按 body
         处理，防止「封面 + 正文同节」时整节误判为封面。
    """
    if _looks_like_body(text_all):
        return "body"
    head = text_all[:HEAD_LOOK]
    if RE_DECL_HINT.search(head):
        return "declaration"
    if RE_COVER_HINT.search(head):
        return "cover"
    return "body"


def leading_front_end(blocks):
    """
    划出文档开头「封面 / 声明页」区的最后一个块号（无该区则返回 0）。

    源稿常把封面与正文放在同一节（.doc 转换后往往只有 1 个 sectPr），
    这时只靠分节判不出封面，封面会被当成正文改写。故按**位置**再划一次：
    从第 1 块起向后扫，遇到第一个**章标题**或**前置部分起始标记**
    （摘要 / Abstract / 目录 / 声明 / 参考文献…）即停，之前的块即封面区。

    注意：返回的只是一段**候选**区间，调用方还须用 section_kind() 确认它
    确实像封面/声明页，才可以把该区的块改判（见 block_kinds）。
    """
    end = 0
    for kind, el, no in blocks:
        if kind == "sectPr":
            continue
        if kind == "tbl":
            end = no
            continue
        t = para_text(el).strip()
        if not t:
            end = no
            continue
        if RE_BODY_START.match(t) or RE_FRONT_START.match(t):
            break
        end = no
    return end


def block_kinds(blocks, sections):
    """
    计算「块号 → 节类型」映射，返回 (no_to_kind, sec_kinds)。

    两条来源合并：
      1) 分节判型（权威）：每节的整节文本交给 section_kind()；
      2) 位置兜底：源稿常把封面与正文放在同一节，第 1 条判不出封面，
         于是再用 leading_front_end() 划出开头的前置区，若该区本身就像
         封面/声明页，则把落在该区、且尚未被判为 cover/declaration 的块改判。

    第 2 条只做「升级」，从不把已是 cover/declaration 的块改回 body；
    且只在候选区确实匹配封面/声明特征时才生效，避免误伤正常正文。
    """
    sec_kinds = [section_kind(range_text(blocks, s, e)) for sec, s, e in sections]
    no_to_kind = {}
    for i, (sec, s, e) in enumerate(sections):
        for n in range(s, e + 1):
            no_to_kind[n] = sec_kinds[i]

    lead_end = leading_front_end(blocks)
    if lead_end > 0:
        lead_kind = section_kind(range_text(blocks, 1, lead_end))
        if lead_kind in ("cover", "declaration"):
            for n in range(1, lead_end + 1):
                if no_to_kind.get(n, "body") == "body":
                    no_to_kind[n] = lead_kind
    return no_to_kind, sec_kinds


def parse_rels(parts, part_name):
    """
    解析指定部件的 .rels，返回 {rId: 目标部件绝对名}。

    Target 通常是相对路径（如 "header6.xml"），此处统一补上部件所在目录，
    否则会查不到页眉页脚部件。External 关系（超链接等）原样返回。
    """
    d, base = os.path.split(part_name)
    rels_name = (d + "/_rels/" + base + ".rels") if d else ("_rels/" + base + ".rels")
    if rels_name not in parts:
        return {}
    root = parse_xml(parts[rels_name])
    out = {}
    for rel in root:
        rid = rel.get("Id")
        tgt = (rel.get("Target") or "").replace("\\", "/")
        if rel.get("TargetMode") == "External" or tgt.startswith(("http:", "https:", "mailto:")):
            out[rid] = tgt
            continue
        if tgt.startswith("/"):
            tgt = tgt.lstrip("/")
        elif d:
            tgt = os.path.normpath(d + "/" + tgt).replace(os.sep, "/")
        out[rid] = tgt
    return out
