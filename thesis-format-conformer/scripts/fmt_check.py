# -*- coding: utf-8 -*-
"""
论文格式检查器 —— 只读诊断，不修改文件

用法：
    python fmt_check.py <论文.docx> [--json 报告路径] [--max-items 40]

输出：
    1. 页面与分节问题
    2. 样式表问题（缺失 / 定义不符）
    3. 段落角色与格式偏差（按类型聚合 + 明细）
    4. 页眉页脚问题

设计原则：只报告偏差，不写入；修复由 fmt_apply.py 完成。
"""
import argparse
import io
import json
import os
import sys
from collections import Counter, OrderedDict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import docx_kit as kit          # noqa: E402
import spec_rules as spec       # noqa: E402

# 参与比对的"核心属性"
CORE_KEYS = ["sz", "eastAsia", "ascii", "bold", "jc", "line", "lineRule",
             "firstLine", "left", "hanging", "outlineLvl", "pageBreakBefore"]

KEY_LABEL = {
    "sz": "字号", "eastAsia": "中文字体", "ascii": "西文字体", "bold": "加粗",
    "jc": "对齐", "line": "行距值", "lineRule": "行距规则", "firstLine": "首行缩进",
    "left": "左缩进", "hanging": "悬挂缩进", "outlineLvl": "大纲级别",
    "pageBreakBefore": "段前分页",
}


def fmt_pt(x):
    if x is None:
        return "未设置"
    for name, v in spec.CN_FONT_SIZE.items():
        if abs(v - x) < 0.01:
            return "%gpt(%s)" % (x, name)
    return "%gpt" % x


def fmt_val(key, v):
    if v is None:
        return "未设置"
    if key == "sz":
        return fmt_pt(v)
    if key in ("firstLine", "left", "hanging"):
        return "%s twips" % v
    if key == "line" and v == 0:
        return "0(最小)"
    if key == "line" and isinstance(v, int) and 1 <= v <= 1530:
        return "%s(%s倍)" % (v, round(v / 240.0, 2))
    if key == "bold":
        return "是" if v else "否"
    if key == "pageBreakBefore":
        return "是" if v else "否"
    return str(v)


def norm_jc(v):
    return {"start": "left", "end": "right"}.get(v, v)


def diff_props(actual, expect, keys=None):
    """返回 [(key, actual, expect)]"""
    out = []
    for k in (keys or CORE_KEYS):
        if k not in expect or expect[k] is None:
            continue
        a = actual.get(k)
        if k == "jc":
            a = norm_jc(a) if a is not None else None
        if a != expect[k]:
            out.append((k, a, expect[k]))
    return out


def check(path, max_items=40):
    parts = kit.load_parts(path)
    doc = kit.parse_xml(parts["word/document.xml"])
    body = doc.find(kit.w("body"))
    styles = kit.StyleTable(kit.parse_xml(parts["word/styles.xml"]))

    report = OrderedDict()
    problems = Counter()

    # ---------- 基本统计 ----------
    n_para = sum(1 for k, _, _ in kit.iter_blocks(body) if k == "p")
    n_tbl = sum(1 for k, _, _ in kit.iter_blocks(body) if k == "tbl")
    sections = kit.collect_sections(body)
    report["overview"] = {
        "file": os.path.basename(path),
        "paragraphs": n_para, "tables": n_tbl, "sections": len(sections),
    }

    # ---------- 1. 页面与分节 ----------
    page_issues = []
    body_section_idx = None
    blocks = list(kit.iter_blocks(body))

    # 建立 段落号→元素 与 段落号→节类型 映射
    para_map = {no: el for kind, el, no in blocks if kind == "p"}
    # 块号→节类型：分节判型 + 位置兜底（源稿常把封面与正文放在同一节）
    no_to_kind, sec_kinds = kit.block_kinds(blocks, sections)

    # 定位"正文起始节"：含第一个章标题、且不属于封面/声明页的节
    first_chapter_no = None
    for kind, el, no in blocks:
        if kind == "p":
            props, sid = kit.para_effective(el, styles)
            role = kit.detect_role(kit.para_text(el), props, styles, sid,
                                   sect_kind=no_to_kind.get(no, "body"))
            if role == "chapter_title":
                first_chapter_no = no
                break
    for i, (sec, s, e) in enumerate(sections, 1):
        if first_chapter_no is not None and s <= first_chapter_no <= e:
            body_section_idx = i
        if sec.pg_w != spec.PAGE_W or sec.pg_h != spec.PAGE_H:
            page_issues.append({
                "section": i, "kind": "纸张尺寸",
                "detail": "当前 %s×%s twips" % (sec.pg_w, sec.pg_h),
                "expect": "%s×%s twips（A4 21.0×29.7cm）" % (spec.PAGE_W, spec.PAGE_H)})
        # 封面节的上边距/页眉距更小，其余节一致
        margins = spec.MARGIN_COVER if sec_kinds[i - 1] == "cover" else spec.MARGIN_BODY
        for k, ev in margins.items():
            av = sec.margins.get(k)
            if av != ev:
                page_issues.append({
                    "section": i, "kind": "页边距 " + k,
                    "detail": "当前 %s" % av, "expect": "%s twips" % ev})
        if sec.doc_grid_pitch != spec.DOC_GRID_LINE_PITCH:
            page_issues.append({
                "section": i, "kind": "文档网格 linePitch",
                "detail": "当前 %s" % sec.doc_grid_pitch,
                "expect": "%s twips" % spec.DOC_GRID_LINE_PITCH})
        # 页码格式
        if body_section_idx and i == body_section_idx:
            if sec.pgnum_fmt not in (None, "decimal") or sec.pgnum_start != 1:
                page_issues.append({
                    "section": i, "kind": "正文页码起始",
                    "detail": "fmt=%s start=%s" % (sec.pgnum_fmt, sec.pgnum_start),
                    "expect": "decimal start=1（正文重新从 1 编号）"})
    report["page_issues"] = page_issues
    problems["page"] = len(page_issues)

    # ---------- 2. 样式表 ----------
    style_issues = []
    for name, sp in spec.STYLE_SPECS.items():
        sid = styles.find_by_name(name)
        if sid is None:
            style_issues.append({"style": name, "kind": "缺失", "detail": "文档中不存在该样式定义"})
            continue
        eff = styles.effective_style_id(sid)
        diffs = diff_props(eff, sp)
        if diffs:
            style_issues.append({
                "style": name, "kind": "定义不符",
                "detail": "; ".join("%s %s→%s" % (KEY_LABEL.get(k, k), fmt_val(k, a), fmt_val(k, e))
                                    for k, a, e in diffs)})
    report["style_issues"] = style_issues
    problems["style"] = len(style_issues)

    # ---------- 3. 段落 ----------
    para_issues = []
    num_issues = []
    role_stat = Counter()
    blank_cnt = 0
    in_ref = False
    ref_started = False
    for kind, el, no in blocks:
        if kind == "tbl":
            continue
        text = kit.para_text(el)
        props, sid = kit.para_effective(el, styles)
        cur_name = kit.style_name_of(styles, sid) or "(无样式)"
        role = kit.detect_role(text, props, styles, sid, in_reference_block=in_ref,
                               sect_kind=no_to_kind.get(no, "body"))
        if role == "blank":
            blank_cnt += 1
            continue
        role_stat[role] += 1

        # 标题编号冲突：规范样式自带自动编号，文本里若仍留手写编号会渲染成
        # 「1.1 1.1 标题」；章标题则相反 —— 只有挂上段落级 numPr 后才会重复。
        stripped = text.strip()
        manual = None
        if role in spec.CONFLICT_HEADING_ROLES:
            m = spec.RE_MANUAL_DECIMAL.match(stripped) or spec.RE_MANUAL_OUTLINE.match(stripped)
            if m:
                manual = m.group(0).strip()
        elif role == "chapter_title":
            ppr = el.find(kit.w("pPr"))
            has_numpr = ppr is not None and ppr.find(kit.w("numPr")) is not None
            m = spec.RE_MANUAL_CHAPTER.match(stripped)
            if has_numpr and m:
                manual = m.group(0).strip()
        if manual:
            num_issues.append({
                "para": no, "role": role, "preview": stripped[:40],
                "detail": "标题文本以手写编号「%s」开头，而规范样式会再自动编号一次，"
                          "Word 中将显示为重复编号" % manual})
        # 可疑合并标题：一段里挤了两条标题（源稿内容问题，需人工确认后改）
        if role in spec.CONFLICT_HEADING_ROLES or role == "chapter_title":
            mh = spec.RE_MERGED_HEADING.match(stripped)
            if mh:
                num_issues.append({
                    "para": no, "role": role, "preview": stripped[:40],
                    "detail": "疑似一段里挤了两条标题（「%s」+「%s」）；请核对后决定保留哪一条，"
                              "或用 --replace-file 显式改写" % (mh.group(1), mh.group(2))})
        if role == "major_title":
            compact = text.strip().replace(" ", "").replace("\u3000", "")
            if compact in ("参考文献", "References"):
                in_ref = True
                ref_started = True
            elif compact in ("致谢", "附录", "Acknowledgement", "Appendix"):
                in_ref = False
        want_style = spec.ROLE_TO_STYLE.get(role)
        if want_style is None:
            continue
        sp = spec.STYLE_SPECS.get(want_style, {})
        issues = []
        # a) 样式不匹配
        if cur_name.strip().lower() != want_style.strip().lower():
            issues.append("样式：当前「%s」→ 应「%s」" % (cur_name, want_style))
        # b) 格式偏差（在已套对样式时，仍检查直接格式覆盖）
        diffs = diff_props(props, sp)
        for k, a, e in diffs:
            issues.append("%s：%s → %s" % (KEY_LABEL.get(k, k), fmt_val(k, a), fmt_val(k, e)))
        if issues:
            para_issues.append({
                "para": no, "role": role, "style": cur_name,
                "preview": text.strip()[:40],
                "issues": issues})
    report["role_stat"] = dict(role_stat)
    report["blank_paragraphs"] = blank_cnt
    report["para_issues_total"] = len(para_issues)
    report["para_issues"] = para_issues[:max_items]
    report["para_issues_truncated"] = max(0, len(para_issues) - max_items)
    problems["paragraph"] = len(para_issues)
    report["num_issues_total"] = len(num_issues)
    report["num_issues"] = num_issues[:max_items]
    report["num_issues_truncated"] = max(0, len(num_issues) - max_items)
    problems["numbering"] = len(num_issues)

    # ---------- 4. 页眉页脚 ----------
    hf_issues = []
    hf_notes = []
    rels = kit.parse_rels(parts, "word/document.xml")
    targets = rels
    last_headers, last_footers = {}, {}
    for i, (sec, s, e) in enumerate(sections, 1):
        # OOXML 规则：本节未定义页眉/页脚时，继承上一节的
        headers = sec.headers or last_headers
        footers = sec.footers or last_footers
        if sec.headers:
            last_headers = sec.headers
        if sec.footers:
            last_footers = sec.footers
        has_page_field = False
        has_styleref = False
        header_texts = []
        for typ, rid in list(headers.items()):
            tgt = targets.get(rid)
            if not tgt or tgt not in parts:
                continue
            root = kit.parse_xml(parts[tgt])
            txt = "".join(t.text or "" for t in root.iter(kit.w("t"))).strip()
            header_texts.append(txt)
            if any("STYLEREF" in (t.text or "") for t in root.iter(kit.w("instrText"))):
                has_styleref = True
        for typ, rid in list(footers.items()):
            tgt = targets.get(rid)
            if not tgt or tgt not in parts:
                continue
            root = kit.parse_xml(parts[tgt])
            if any("PAGE" in (t.text or "") for t in root.iter(kit.w("instrText"))):
                has_page_field = True
        is_body = (body_section_idx == i)
        if is_body:
            hdr_joined = " / ".join(x for x in header_texts if x)
            school_mode = bool(spec.SCHOOL_NAME) and spec.SCHOOL_NAME in hdr_joined
            if not headers:
                hf_issues.append({"section": i, "kind": "页眉缺失", "detail": "正文节未定义页眉"})
            elif not has_styleref and not school_mode:
                hf_issues.append({"section": i, "kind": "页眉缺少 STYLEREF 域",
                                  "detail": "奇数页页眉应含 STYLEREF \"标题 1\" 域（自动显示当前章名）；当前页眉：%s"
                                            % hdr_joined})
            elif not has_styleref and school_mode:
                hf_notes.append("节 %d：使用统一校名页眉「%s」（奇偶页相同，不含章名域），"
                                "属 --school-header 模式的预期结果。" % (i, hdr_joined))
            if not has_page_field:
                hf_issues.append({"section": i, "kind": "页脚缺少页码域",
                                  "detail": "正文节页脚应含 PAGE 域（居中）"})
    report["header_footer_issues"] = hf_issues
    report["header_footer_notes"] = hf_notes
    problems["header_footer"] = len(hf_issues)

    report["problem_summary"] = dict(problems)
    report["problem_total"] = sum(problems.values())
    return report


def render(report):
    out = io.StringIO()
    ov = report["overview"]
    out.write("=" * 78 + "\n")
    out.write("论文格式检查报告\n")
    out.write("=" * 78 + "\n")
    out.write("文件: %s\n" % ov["file"])
    out.write("结构: 段落/块 %d 个 | 表格 %d 个 | 分节 %d 个\n" % (
        ov["paragraphs"], ov["tables"], ov["sections"]))
    ps = report["problem_summary"]
    out.write("问题: 合计 %d 处  [页面 %d | 样式 %d | 段落 %d | 编号 %d | 页眉页脚 %d]\n\n" % (
        report["problem_total"], ps["page"], ps["style"], ps["paragraph"],
        ps["numbering"], ps["header_footer"]))

    out.write("--- 一、页面与分节 ---\n")
    if not report["page_issues"]:
        out.write("  符合规范。\n")
    for it in report["page_issues"][:30]:
        out.write("  [节 %s] %s：%s（应为 %s）\n" % (it["section"], it["kind"], it["detail"], it["expect"]))

    out.write("\n--- 二、样式表 ---\n")
    if not report["style_issues"]:
        out.write("  规范样式齐备且定义一致。\n")
    for it in report["style_issues"]:
        out.write("  [%s] %s：%s\n" % (it["kind"], it["style"], it["detail"]))

    out.write("\n--- 三、段落角色分布 ---\n")
    for role, cnt in sorted(report["role_stat"].items(), key=lambda x: -x[1]):
        out.write("  %-22s %d\n" % (role, cnt))
    out.write("  空段落: %d\n" % report["blank_paragraphs"])

    out.write("\n--- 四、段落格式偏差（共 %d 处） ---\n" % report["para_issues_total"])
    for it in report["para_issues"]:
        out.write("  [段 %d] 角色=%s 当前样式=「%s」  %r\n" % (it["para"], it["role"], it["style"], it["preview"]))
        for s in it["issues"]:
            out.write("        - %s\n" % s)
    if report["para_issues_truncated"]:
        out.write("  ...另有 %d 处未列出（用 --max-items 调整）\n" % report["para_issues_truncated"])

    out.write("\n--- 五、页眉页脚 ---\n")
    if not report["header_footer_issues"]:
        out.write("  符合规范。\n")
    for it in report["header_footer_issues"]:
        out.write("  [节 %s] %s：%s\n" % (it["section"], it["kind"], it["detail"]))
    for note in report.get("header_footer_notes", []):
        out.write("  （说明）%s\n" % note)

    out.write("\n--- 六、标题编号冲突（共 %d 处）---\n" % report["num_issues_total"])
    if not report["num_issues_total"]:
        out.write("  未发现手写编号与自动编号重复。\n")
    for it in report["num_issues"]:
        out.write("  [段 %d] 角色=%s  %r\n" % (it["para"], it["role"], it["preview"]))
        out.write("        - %s\n" % it["detail"])
    if report["num_issues_truncated"]:
        out.write("  ...另有 %d 处未列出（用 --max-items 调整）\n" % report["num_issues_truncated"])
    out.write("\n")
    return out.getvalue()


def main():
    ap = argparse.ArgumentParser(description="论文格式检查（只读）")
    ap.add_argument("docx", help="待检查的 .docx 路径")
    ap.add_argument("--json", help="同时输出 JSON 报告的路径")
    ap.add_argument("--max-items", type=int, default=40, help="段落明细最多列出条数")
    args = ap.parse_args()

    report = check(args.docx, max_items=args.max_items)
    text = render(report)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print(text)
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        print("JSON 报告已写入: %s" % args.json)


if __name__ == "__main__":
    main()
