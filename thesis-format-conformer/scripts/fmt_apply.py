# -*- coding: utf-8 -*-
"""
论文格式规范化应用器

用法：
    python fmt_apply.py <输入.docx> [-o 输出.docx] [--dry-run] [--report 报告.txt]

默认输出 <输入文件名>-已规范化.docx，绝不覆盖原文件。

修复内容（三块）：
    1. 页面与分节 —— A4 纸张、规范页边距、文档网格、页码格式（正文重新从 1 起）
    2. 样式表     —— 按样式名同步规范样式定义（保留目标文档已有 ID，缺失则注入）
    3. 段落       —— 按语义角色套用规范样式，并清除与规范冲突的直接格式

典型用法（在 WorkBuddy 中）：
    先跑 fmt_check.py 得到偏差报告 → 跑 fmt_apply.py 修复 → 再跑 fmt_check.py 复核。
"""
import argparse
import os
import re
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import docx_kit as kit          # noqa: E402
import spec_rules as spec       # noqa: E402

W = kit.W_NS
HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")
REF_STYLES = os.path.join(ASSETS, "reference-styles.xml")
REF_NUMBERING = os.path.join(ASSETS, "reference-numbering.xml")

# 需要同步到目标文档的样式白名单（论文真正用到的），其 basedOn/link 依赖会自动补齐
NEED_STYLE_NAMES = [
    # 正文与标题
    "Normal", "我的正文", "我的正文 Char",
    "heading 1", "heading 2", "heading 3", "heading 4",
    "标题 1 字符", "标题 2 字符", "标题 3 字符", "标题 4 字符",
    # 篇级大标题
    "我的标题", "论文标题", "论文题目", "内容摘要",
    # 图表
    "图标题", "图所在段落", "表标题", "表格内文字", "图表说明", "图注释",
    "表格说明", "公式", "公式段落", "caption",
    # 参考文献与列表
    "文献列表", "List Paragraph", "列表段落1", "书目1",
    # 表格
    "三线表", "Table Grid", "Normal Table",
    # 页眉页脚与脚注
    "header", "footer", "脚注文本", "footnote text",
    "page number", "footnote reference", "Hyperlink",
    "页眉 字符", "页脚 字符",
    "我的正文倾斜", "我的正文-倾斜英文", "我的正文-倾斜英文 Char Char",
    # 目录
    "toc 1", "toc 2", "toc 3", "toc 4", "toc 5", "toc 6", "toc 7", "toc 8", "toc 9",
]

# pPr 子元素标准顺序（OOXML CT_PPr），插入时必须遵守
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]
# 段落级直接格式中需清除的项（交由样式决定）
# 注意：**不含 numPr** —— 规范里「第X章」正是靠段落级 numPr 生成的，
# 一并清除会让章标题彻底失去编号（此坑已在厦大范文上实测复现）。
# numPr 改由 normalize_heading_number() 按层级显式设置。
CLEAR_PPR = ["jc", "ind", "spacing", "outlineLvl", "pageBreakBefore", "keepNext"]
# run 级直接格式中需清除的项（字体字号交由样式决定；保留 i/u/vertAlign/color 等强调效果）
CLEAR_RPR = ["rFonts", "sz", "szCs", "b", "bCs"]

# 标题层级 -> 多级编号的 ilvl（章=0 靠段落级 numPr；其余靠样式级 numPr）
HEADING_ILVL = {
    "chapter_title": 0,
    "section_title": 1,
    "subsection_title": 2,
    "subsubsection_title": 3,
}
# 标题文本里的手写编号前缀（剥离后由自动编号生成）
RE_CHAP_NUM_PREFIX = re.compile(
    r"^\s*第\s*[一二三四五六七八九十百零〇0-9]+\s*[章篇]\s*[、:：.．]?\s*")
RE_DEC_NUM_PREFIX = re.compile(r"^\s*\d+(?:\.\d+)*[\s、:：.．]*")
# 不该被当作编号剥离的后续字符（防御性：正文句 "2023年3月…"）
RE_NOT_NUMBER_SUFFIX = re.compile(r"^(年|月|日|%|％|年代|年度)")


def _insert_ordered(parent, el, tag):
    """按 OOXML 规定的子元素顺序插入。"""
    idx = PPR_ORDER.index(tag)
    for child in parent:
        ctag = child.tag.split("}")[-1]
        if ctag in PPR_ORDER and PPR_ORDER.index(ctag) > idx:
            child.addprevious(el)
            return
    parent.append(el)


def _drop_children(parent, tags):
    if parent is None:
        return 0
    n = 0
    for tag in tags:
        for el in parent.findall(kit.w(tag)):
            parent.remove(el)
            n += 1
    return n


# ---------------------------------------------------------------------------
# 1. 样式同步
# ---------------------------------------------------------------------------
def sync_styles(tgt_root, ref_root, log):
    """把规范样式同步进目标文档，返回 {样式名小写: 最终 styleId}。"""
    # 目标文档现有样式
    tgt_by_name = {}
    tgt_ids = set()
    for st in tgt_root.findall(kit.w("style")):
        sid = st.get(kit.w("styleId"))
        tgt_ids.add(sid)
        nm = st.find(kit.w("name"))
        if nm is not None and nm.get(kit.w("val")):
            tgt_by_name[nm.get(kit.w("val")).strip().lower()] = sid

    # 规范样式
    ref_by_id = {}
    ref_by_name = {}
    for st in ref_root.findall(kit.w("style")):
        sid = st.get(kit.w("styleId"))
        ref_by_id[sid] = st
        nm = st.find(kit.w("name"))
        if nm is not None and nm.get(kit.w("val")):
            ref_by_name[nm.get(kit.w("val")).strip().lower()] = st

    # 白名单 + 依赖闭包（basedOn / link / next 指向的样式一并纳入）
    wanted_ids = set()
    stack = []
    for name in NEED_STYLE_NAMES:
        st = ref_by_name.get(name.strip().lower())
        if st is not None:
            stack.append(st.get(kit.w("styleId")))
    while stack:
        sid = stack.pop()
        if sid in wanted_ids or sid not in ref_by_id:
            continue
        wanted_ids.add(sid)
        st = ref_by_id[sid]
        for tag in ("basedOn", "link", "next"):
            child = st.find(kit.w(tag))
            if child is not None and child.get(kit.w("val")):
                stack.append(child.get(kit.w("val")))

    # 先决定每个规范样式的最终 ID（同名者沿用目标 ID）
    id_map = {}     # ref_id -> final_id
    name2final = {}
    used_ids = set(tgt_ids)
    for rid in sorted(wanted_ids):
        st = ref_by_id[rid]
        nm = st.find(kit.w("name"))
        name = (nm.get(kit.w("val")) if nm is not None else "").strip()
        key = name.lower()
        if key and key in tgt_by_name:
            final = tgt_by_name[key]
        else:
            final = rid
            base, i = rid, 1
            while final in used_ids:
                final = "%s_f%d" % (base, i)
                i += 1
            used_ids.add(final)
        id_map[rid] = final
        if key:
            name2final[key] = final

    # 写入（替换同名样式 / 追加新样式）
    replaced = added = 0
    for rid in sorted(wanted_ids):
        st = ref_by_id[rid]
        nm = st.find(kit.w("name"))
        name = (nm.get(kit.w("val")) if nm is not None else "").strip()
        new_el = etree.fromstring(etree.tostring(st))
        new_el.set(kit.w("styleId"), id_map[rid])
        for tag in ("basedOn", "link", "next"):
            child = new_el.find(kit.w(tag))
            if child is not None:
                v = child.get(kit.w("val"))
                if v in id_map:
                    child.set(kit.w("val"), id_map[v])
        final = id_map[rid]
        cur = None
        for s in tgt_root.findall(kit.w("style")):
            if s.get(kit.w("styleId")) == final:
                cur = s
                break
        if cur is not None:
            tgt_root.replace(cur, new_el)
            replaced += 1
        else:
            tgt_root.append(new_el)
            added += 1
    log.append("样式同步：替换 %d 个、新增 %d 个（共处理 %d 个规范样式）"
               % (replaced, added, len(wanted_ids)))
    return name2final


# ---------------------------------------------------------------------------
# 2. 段落修复
# ---------------------------------------------------------------------------
def set_para_style(p, style_id, clear_direct=True):
    """设置段落样式并清除与规范冲突的直接格式。返回清除项数。"""
    ppr = p.find(kit.w("pPr"))
    if ppr is None:
        ppr = etree.Element(kit.w("pPr"))
        p.insert(0, ppr)
    cleared = 0
    if clear_direct:
        cleared += _drop_children(ppr, CLEAR_PPR)
        inner_rpr = ppr.find(kit.w("rPr"))
        if inner_rpr is not None:
            cleared += _drop_children(inner_rpr, CLEAR_RPR)
    ps = ppr.find(kit.w("pStyle"))
    if ps is None:
        ps = etree.Element(kit.w("pStyle"))
        _insert_ordered(ppr, ps, "pStyle")
    if ps.get(kit.w("val")) != style_id:
        ps.set(kit.w("val"), style_id)
        cleared += 1
    # run 级字体字号清理
    for r in p.findall(kit.w("r")):
        rpr = r.find(kit.w("rPr"))
        cleared += _drop_children(rpr, CLEAR_RPR)
    return cleared


def set_para_numpr(p, num_id, ilvl):
    """设置（或清除）段落级编号。num_id=None 表示清除。"""
    ppr = p.find(kit.w("pPr"))
    if ppr is None:
        ppr = etree.Element(kit.w("pPr"))
        p.insert(0, ppr)
    for el in ppr.findall(kit.w("numPr")):
        ppr.remove(el)
    if num_id is None:
        return
    numpr = etree.Element(kit.w("numPr"))
    il = etree.SubElement(numpr, kit.w("ilvl"))
    il.set(kit.w("val"), str(ilvl))
    nid = etree.SubElement(numpr, kit.w("numId"))
    nid.set(kit.w("val"), str(num_id))
    _insert_ordered(ppr, numpr, "numPr")


def set_para_ind(p, left=None, firstLine=None, hanging=None):
    """
    设置段落级缩进（覆盖样式/编号自带的缩进）。None 表示不写该项。

    **章标题必须用 left=0 覆盖编号自带缩进**：
    多级编号 abstractNum 的 ilvl=0 定义为 `ind left=4819 firstLine=0`（约 8.5cm），
    范文靠段落级 `w:ind w:left="0"` 把它抵消掉；缺失时章标题会被推到页面右侧、
    长标题（如「综合监管目的和银行风险」）当场折行——已实测复现。
    """
    ppr = p.find(kit.w("pPr"))
    if ppr is None:
        ppr = etree.Element(kit.w("pPr"))
        p.insert(0, ppr)
    old = ppr.find(kit.w("ind"))
    if old is not None:
        ppr.remove(old)
    vals = {}
    if left is not None:
        vals["left"] = str(left)
    if firstLine is not None:
        vals["firstLine"] = str(firstLine)
    if hanging is not None:
        vals["hanging"] = str(hanging)
    if not vals:
        return
    ind = etree.Element(kit.w("ind"))
    for k, v in vals.items():
        ind.set(kit.w(k), v)
    _insert_ordered(ppr, ind, "ind")


# ---------------------------------------------------------------------------
# 显式文本替换（不做启发式：只在用户给定「原文 → 新文」时改）
# ---------------------------------------------------------------------------
def norm_text(s):
    """归一化：去掉所有空白（含全角），用于宽松匹配整段文本。"""
    return re.sub(r"[\s\u3000]+", "", s or "")


def strip_outline_number(s):
    """去掉段首可能存在的章/节编号前缀，便于替换表不必写编号。"""
    for rx in (RE_CHAP_NUM_PREFIX, RE_DEC_NUM_PREFIX):
        m = rx.match(s)
        if m:
            return s[m.end():]
    return s


def load_replacements(path):
    """
    读取替换表：每行「原文 <TAB> 新文」，# 开头为注释，空行忽略。
    匹配时忽略空白差异、也忽略段首编号，因此不必写「1 」「2.1 」这类前缀。
    """
    table = []
    with open(path, "r", encoding="utf-8-sig") as f:
        for ln, line in enumerate(f, 1):
            line = line.rstrip("\n").rstrip("\r")
            if not line.strip() or line.lstrip().startswith("#"):
                continue
            if "\t" in line:
                old, new = line.split("\t", 1)
            elif "=>" in line:
                old, new = line.split("=>", 1)
            else:
                raise SystemExit("替换表第 %d 行缺少分隔符（Tab 或 =>）：%s" % (ln, line))
            table.append((old.strip(), new.strip()))
    return table


def set_para_text(p, text):
    """把整段文本换成 text：写入第一个 w:t，其余 w:t 清空（保留首个 run 的格式）。"""
    ts = list(p.iter(kit.w("t")))
    if not ts:
        r = etree.SubElement(p, kit.w("r"))
        t = etree.SubElement(r, kit.w("t"))
        ts = [t]
    ts[0].text = text
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for t in ts[1:]:
        t.text = ""
    return len(ts)


def build_replace_index(table):
    """把替换表编成 {归一化原文: 新文}。同时登记「去掉段首编号」的键，便于宽松命中。"""
    idx = {}
    for old, new in table or []:
        idx[norm_text(old)] = new
        idx[norm_text(strip_outline_number(old))] = new
    return idx


def lookup_replacement(idx, text):
    if not idx:
        return None
    key = norm_text(text)
    if key in idx:
        return idx[key]
    return idx.get(norm_text(strip_outline_number(text)))


def prune_empty_runs(p):
    """
    删除剥离编号后残留的空 run / 空 w:t。

    剥离「1.1 」「第二章 」时，被删字符所在的 run 会留下空的 <w:t/>。
    空文本不渲染（Word 的 Range.Text 也读不到），但会让 XML 难读；
    且部分编辑器在「显示编辑标记」下会把它画成一个小方块，
    容易被误当成文档里的乱码。故一并清掉。
    """
    removed = 0
    for r in list(p.findall(kit.w("r"))):
        for t in list(r.findall(kit.w("t"))):
            if not (t.text or ""):
                r.remove(t)
                removed += 1
        rest = [c for c in r if etree.QName(c).localname != "rPr"]
        if not rest:
            p.remove(r)
            removed += 1
    return removed


def normalize_heading_number(p, role, heading_num_id):
    """
    标题编号归一：剥离文本里的手写编号，改由自动编号生成。

    规范机制（与厦大范文一致）——
      章标题：文本只剩标题，编号由**段落级** numPr(numId, ilvl=0) 生成「第X章」，
              并须写 `w:ind left=0` 抵消编号自带的 8.5cm 左缩进（否则长标题折行）；
      节及以下：编号由**样式级** numPr 生成「1.1 / 1.1.1」，段落级不得再有 numPr 覆盖。

    若不剥离，套上带自动编号的 heading 样式后会渲染成「1.1 1.1 选题背景」。
    返回被剥离的前缀文本（供日志记录）。
    """
    text = kit.para_text(p)
    cut = 0
    if role == "chapter_title":
        m = RE_CHAP_NUM_PREFIX.match(text)
        if m:
            cut = m.end()
        set_para_numpr(p, heading_num_id, HEADING_ILVL[role])
        set_para_ind(p, left=0, firstLine=0)
    else:
        m = RE_DEC_NUM_PREFIX.match(text)
        if m and not RE_NOT_NUMBER_SUFFIX.match(text[m.end():]):
            cut = m.end()
        set_para_numpr(p, None, None)
    if cut:
        kit.strip_prefix_chars(p, cut)
        prune_empty_runs(p)
    return text[:cut]


# ---------------------------------------------------------------------------
# 3. 页面与分节
# ---------------------------------------------------------------------------
def set_pg(section, w_twips=None, h_twips=None, margins=None, pgnum=None, pitch=None):
    """按需设置页面尺寸 / 页边距 / 页码格式 / 文档网格。"""
    el = section.el
    if w_twips and h_twips:
        pg = el.find(kit.w("pgSz"))
        if pg is None:
            pg = etree.SubElement(el, kit.w("pgSz"))
        pg.set(kit.w("w"), str(w_twips))
        pg.set(kit.w("h"), str(h_twips))
    if margins:
        mar = el.find(kit.w("pgMar"))
        if mar is None:
            mar = etree.SubElement(el, kit.w("pgMar"))
        for k, v in margins.items():
            mar.set(kit.w(k), str(v))
    if pgnum is not None:
        pn = el.find(kit.w("pgNumType"))
        if pn is None:
            pn = etree.SubElement(el, kit.w("pgNumType"))
        fmt, start = pgnum
        if fmt:
            pn.set(kit.w("fmt"), fmt)
        elif pn.get(kit.w("fmt")):
            del pn.attrib[kit.w("fmt")]
        if start is not None:
            pn.set(kit.w("start"), str(start))
    if pitch is not None:
        dg = el.find(kit.w("docGrid"))
        if dg is None:
            dg = etree.SubElement(el, kit.w("docGrid"))
        dg.set(kit.w("linePitch"), str(pitch))


# ---------------------------------------------------------------------------
# 4. 编号定义
# ---------------------------------------------------------------------------
def sync_numbering(tgt_num, log):
    """
    把规范的多级标题编号（absId=11）与参考文献编号（absId=10）同步进目标文档，
    并确保它们在目标文档里各有一个可用的 numId。

    与旧版的关键差异：**不再假定目标文档的 numId=1 就是标题编号**。
    旧写法只在 numId 缺失时追加，若目标文档恰好已把 numId=1 用在别处，
    标题编号就会指向错误的定义。这里改为按 abstractNumId 反查，
    找不到才用空闲 numId 新建，绝不覆盖目标文档已有的映射。

    返回标题多级编号的 numId（供章标题的段落级 numPr 引用）。
    """
    ref_num = kit.parse_xml(open(REF_NUMBERING, "rb").read())
    ref_abs = {a.get(kit.w("abstractNumId")): a for a in ref_num.findall(kit.w("abstractNum"))}
    ref_nums = {n.get(kit.w("numId")): n for n in ref_num.findall(kit.w("num"))}
    added = 0

    for aid in ("10", "11"):
        if aid not in ref_abs:
            continue
        new_el = etree.fromstring(etree.tostring(ref_abs[aid]))
        cur = None
        for a in tgt_num.findall(kit.w("abstractNum")):
            if a.get(kit.w("abstractNumId")) == aid:
                cur = a
                break
        if cur is not None:
            tgt_num.replace(cur, new_el)
        else:
            first_num = tgt_num.find(kit.w("num"))
            if first_num is not None:
                first_num.addprevious(new_el)
            else:
                tgt_num.append(new_el)
            added += 1

    def ensure_num(abstract_id, prefer_ids):
        for n in tgt_num.findall(kit.w("num")):
            a = n.find(kit.w("abstractNumId"))
            if a is not None and a.get(kit.w("val")) == abstract_id:
                return n.get(kit.w("numId"))
        taken = {n.get(kit.w("numId")) for n in tgt_num.findall(kit.w("num"))}
        for pid in prefer_ids:
            if pid not in taken and pid in ref_nums:
                tgt_num.append(etree.fromstring(etree.tostring(ref_nums[pid])))
                return pid
        nid = 1
        while str(nid) in taken:
            nid += 1
        el = etree.Element(kit.w("num"))
        el.set(kit.w("numId"), str(nid))
        sub = etree.SubElement(el, kit.w("abstractNumId"))
        sub.set(kit.w("val"), abstract_id)
        tgt_num.append(el)
        return str(nid)

    heading_num = ensure_num("11", ["1"])
    item_num = ensure_num("10", ["12"])
    log.append("编号定义：多级标题 numId=%s、参考文献 numId=%s（新增 abstractNum %d 项）"
               % (heading_num, item_num, added))
    return heading_num


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def apply_format(src, out, dry_run=False, log=None, replacements=None):
    log = log if log is not None else []
    parts = kit.load_parts(src)

    doc_root = kit.parse_xml(parts["word/document.xml"])
    body = doc_root.find(kit.w("body"))
    styles_root = kit.parse_xml(parts["word/styles.xml"])
    styles = kit.StyleTable(styles_root)
    ref_root = kit.parse_xml(open(REF_STYLES, "rb").read())

    # ---- 先做一轮角色扫描（同时确定节的类型） ----
    blocks = list(kit.iter_blocks(body))
    para_map = {no: el for kind, el, no in blocks if kind == "p"}
    sections = kit.collect_sections(body)
    # 块号→节类型（cover / declaration / body）。两条来源合并：分节判型 + 位置兜底，
    # 后者专门解决「源稿把封面与正文放在同一节」的情形（否则封面会被当正文改写）。
    no_to_kind, sec_kinds = kit.block_kinds(blocks, sections)
    log.append("文档结构：%d 个段落块 / %d 个表格 / %d 个分节（%s）"
               % (sum(1 for k, _, _ in blocks if k == "p"),
                  sum(1 for k, _, _ in blocks if k == "tbl"),
                  len(sections), ", ".join(sec_kinds)))
    protected = sorted(n for n, k in no_to_kind.items() if k in ("cover", "declaration"))
    if protected:
        log.append("封面/声明区保护：块 %d..%d 共 %d 个块不套正文规范（自由排版）"
                   % (protected[0], protected[-1], len(protected)))

    # ---- 0. 文本替换表（编成索引，逐段按【原文】命中后再改写） ----
    repl_idx = build_replace_index(replacements)
    if repl_idx:
        log.append("文本替换表：%d 条" % len(replacements))

    # ---- 1. 样式同步 ----
    name2id = sync_styles(styles_root, ref_root, log)

    def sid(name):
        return name2id.get(name.strip().lower())

    # ---- 1.5 编号定义（必须先于段落处理，段落编号要引用这里的 numId） ----
    tgt_num = None
    heading_num_id = None
    if "word/numbering.xml" in parts:
        tgt_num = kit.parse_xml(parts["word/numbering.xml"])
        if not dry_run:
            heading_num_id = sync_numbering(tgt_num, log)
        else:
            heading_num_id = "1"   # dry-run 仅用于统计，编号不会写入

    # ---- 2. 段落修复 ----
    role_cnt = {}
    changed = 0
    cleared_total = 0
    retitled = []
    replaced = []
    in_ref = False
    for kind, el, no in blocks:
        if kind != "p":
            continue
        text = kit.para_text(el)
        props, cur_sid = kit.para_effective(el, styles)
        role = kit.detect_role(text, props, styles, cur_sid, in_reference_block=in_ref,
                               sect_kind=no_to_kind.get(no, "body"))
        # 显式文本替换：角色已按【原文】判定完毕，此处再改写文本。
        # 顺序很关键——源稿的提纲小标题靠「1 xxx」这类段首编号识别，
        # 若先替换（新标题不带编号）再判定角色，该段会被降级成正文（已实测踩过）。
        new_text = lookup_replacement(repl_idx, text)
        if new_text is not None:
            set_para_text(el, new_text)
            replaced.append((no, text.strip(), new_text))
            props, cur_sid = kit.para_effective(el, styles)
            text = kit.para_text(el)
        if role == "major_title":
            compact = text.strip().replace(" ", "").replace("\u3000", "")
            if compact in ("参考文献", "References"):
                in_ref = True
            elif compact in ("致谢", "附录", "Acknowledgement", "Appendix"):
                in_ref = False
        want = spec.ROLE_TO_STYLE.get(role)
        if not want:
            continue
        target_id = sid(want)
        if not target_id:
            continue
        role_cnt[role] = role_cnt.get(role, 0) + 1
        if dry_run:
            changed += 1
            continue
        cleared = set_para_style(el, target_id, clear_direct=True)
        # 标题编号归一：剥离手写编号，改由自动编号生成（否则与样式自带编号重复）
        if role in HEADING_ILVL:
            prefix = normalize_heading_number(el, role, heading_num_id)
            if prefix:
                retitled.append((no, role, prefix, kit.para_text(el).strip()))
        cleared_total += cleared
        if cleared:
            changed += 1

    log.append("段落规范化：%d 段套用规范样式（清除直接格式 %d 项）" % (changed, cleared_total))
    for r, c in sorted(role_cnt.items(), key=lambda x: -x[1]):
        log.append("    %-20s %d" % (r, c))
    if replaced:
        log.append("文本替换：命中 %d 处" % len(replaced))
        for no, old, new in replaced:
            log.append("    段%-4d %r → %r" % (no, old[:36], new[:36]))
    if retitled:
        log.append("  标题编号归一：%d 处剥离手写编号，改由自动编号生成" % len(retitled))
        for no, role, prefix, rest in retitled:
            log.append("    段%-4d %-18s 剥离 %-12r → %s" % (no, role, prefix, rest[:30]))

    # ---- 3. 页面与分节 ----
    first_chapter_no = None
    for kind, el, no in blocks:
        if kind != "p":
            continue
        props, cur_sid = kit.para_effective(el, styles)
        if kit.detect_role(kit.para_text(el), props, styles, cur_sid,
                           sect_kind=no_to_kind.get(no, "body")) == "chapter_title":
            first_chapter_no = no
            break

    for i, (sec, s, e) in enumerate(sections, 1):
        kind_i = sec_kinds[i - 1]
        margins = spec.MARGIN_COVER if kind_i == "cover" else spec.MARGIN_BODY
        pgnum = None
        if kind_i in ("cover", "declaration"):
            pgnum = ("", None)          # 不设页码格式
        elif first_chapter_no is not None and s <= first_chapter_no <= e:
            pgnum = (spec.PGNUM_BODY, 1)  # 正文：阿拉伯数字，重新从 1 起
        else:
            pgnum = (spec.PGNUM_FRONT, 1)  # 前置部分：大写罗马数字
        if not dry_run:
            set_pg(sec, spec.PAGE_W, spec.PAGE_H, margins, pgnum, spec.DOC_GRID_LINE_PITCH)
    log.append("页面与分节：%d 节已标准化（A4 / 页边距 / 页码格式 / 文档网格）" % len(sections))

    # ---- 4. 编号定义已提前至 1.5 处理 ----

    # ---- 写回 ----
    if not dry_run:
        out_parts = dict(parts)
        out_parts["word/styles.xml"] = kit.serialize_xml(styles_root)
        out_parts["word/document.xml"] = kit.serialize_xml(doc_root)
        if tgt_num is not None:
            out_parts["word/numbering.xml"] = kit.serialize_xml(tgt_num)
        with zipfile.ZipFile(src, "r") as zin, \
                zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = out_parts.get(item.filename, zin.read(item.filename))
                zout.writestr(item.filename, data)
        log.append("已输出：%s" % out)
    else:
        log.append("（dry-run：未写入任何文件）")
    return log


def main():
    ap = argparse.ArgumentParser(description="论文格式规范化（默认不覆盖原文件）")
    ap.add_argument("docx", help="输入 .docx")
    ap.add_argument("-o", "--out", help="输出 .docx（默认：<原名>-已规范化.docx）")
    ap.add_argument("--dry-run", action="store_true", help="只统计不写文件")
    ap.add_argument("--report", help="把变更日志写入该文本文件")
    ap.add_argument("--replace-file", default=None,
                    help="文本替换表（每行「原文<Tab>新文」），用于显式修正源稿里的错标题")
    args = ap.parse_args()

    src = args.docx
    out = args.out or os.path.splitext(src)[0] + "-已规范化.docx"
    if os.path.abspath(out) == os.path.abspath(src):
        print("拒绝覆盖原文件，请用 -o 指定其他输出路径。")
        return 2
    repl = load_replacements(args.replace_file) if args.replace_file else None
    log = apply_format(src, out, dry_run=args.dry_run, replacements=repl)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print("=" * 70)
    print("论文格式规范化报告")
    print("=" * 70)
    for line in log:
        print(line)
    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            f.write("\n".join(log))
        print("报告已写入：%s" % args.report)
    return 0


if __name__ == "__main__":
    sys.exit(main())
