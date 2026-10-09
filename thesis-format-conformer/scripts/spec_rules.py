# -*- coding: utf-8 -*-
"""
厦门大学硕士学位论文格式规范 —— 规则常量

规范来源：用户提供的样例论文
    《注册制下价格笼子机制对股票定价效率的影响》
（docProps/app.xml 记录 Template = "研究生论文模板"，Pages = 66）

本模块是 fmt_check.py / fmt_apply.py 共用的"规范宪法"。
所有数值均从样例文档的 styles.xml / document.xml / numbering.xml 实测提取，
单位为 Word 原生单位：
    - 长度：twips（1 cm = 567 twips，1 pt = 20 twips）
    - 字号：half-point（sz="24" 即 12pt）
    - 行距：lineRule="auto" 时 line=360 表示 1.5 倍行距（240 = 单倍）
"""
import re

# ============================================================================
# 1. 字号换算表（中文字号 → pt）
# ============================================================================
CN_FONT_SIZE = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5,
}

# ============================================================================
# 2. 页面设置（分节）
# ============================================================================
# A4：21.0 cm × 29.7 cm
PAGE_W = 11906          # twips（= 21.0 cm）
PAGE_H = 16838          # twips（= 29.7 cm）

# 正文/大多数节的页边距
MARGIN_BODY = {
    "top": 1701,        # 3.0 cm
    "right": 1797,      # 3.17 cm
    "bottom": 1440,     # 2.54 cm
    "left": 1797,       # 3.17 cm
    "header": 1134,     # 2.0 cm
    "footer": 992,      # 1.75 cm
    "gutter": 0,
}
# 封面节的页边距（上边距与页眉距更小）
MARGIN_COVER = {
    "top": 1440,        # 2.54 cm
    "right": 1797,
    "bottom": 1440,
    "left": 1797,
    "header": 851,      # 1.5 cm
    "footer": 992,
    "gutter": 0,
}

DOC_GRID_LINE_PITCH = 312   # 文档网格行距基准（twips）
DEFAULT_TAB_STOP = 420      # 默认制表位（twips）= 2 字符

# 页码格式：每节独立编号
# 说明：pgNumType 的 fmt 取值 upperRoman（大写罗马）/ decimal（阿拉伯，缺省）
PGNUM_FRONT = "upperRoman"  # 声明页 / 摘要 / 目录
PGNUM_BODY = "decimal"      # 正文起（重新从 1 开始）

# ---- 学校品牌（封面标志与页眉校名）----
# 规范基线的学校。换学校时只改这三项（或改用命令行参数 --logo / --school-header）。
SCHOOL_NAME = "厦门大学"
SCHOOL_HEADER_TEXT = "厦门大学本科生毕业论文"   # 统一校名页眉的文本
SCHOOL_COVER_LOGO = "assets/xiamen-university-logo.png"  # 封面官方标志（相对 skill 根目录）
SCHOOL_LOGO_WIDTH_PT = 300.0    # 封面标志宽度（pt）；高度按原图比例自动算
SCHOOL_LOGO_TOP_PT = 28.0       # 封面标志距页顶（pt）

# ---- 页眉页脚规范（由 fmt_hf.py 使用）----
# 三节结构：封面（无页眉页脚、无页码）→ 前置（摘要/Abstract/目录，罗马页码）→ 正文（阿拉伯页码重启）
HF_FRONT_TITLES = ("摘要", "abstract", "目录", "contents", "声明", "原创性声明")  # 前置部分的篇级标题
HEADER_STYLEREF = "标题 1"        # 奇数页页眉引用的样式名（Word 内置「标题 1」的本地化名）
HEADER_EVEN_IS_TITLE = True       # 偶数页页眉显示论文题目
HDRFTR_SIZE = 21                  # 页眉页脚字号（半磅），21 = 10.5pt 五号

# ============================================================================
# 3. 样式规范（按样式名索引，不依赖 styleId —— styleId 是 Word 内部随机 ID）
# ============================================================================
# 每条规范的字段与"有效格式"抽取结构一致，None 表示不约束：
#   jc         对齐：both/center/left/right/distribute
#   firstLine  首行缩进（twips）
#   left       左缩进（twips）
#   hanging    悬挂缩进（twips）
#   line       行距数值；lineRule=auto 时 240=单倍、360=1.5倍
#   lineRule   auto/exact/atLeast
#   before/after 段前/段后间距（twips）
#   eastAsia   中文字体
#   ascii      西文字体
#   sz         字号（pt）
#   bold       加粗
#   outlineLvl 大纲级别
#   pageBreakBefore 段前分页

STYLE_SPECS = {
    # 说明：中文学位论文的格式审核以「中文字体 + 字号 + 间距 + 缩进 + 大纲级别」为准；
    # 西文字体（Times New Roman / Arial）随文档默认或样式定义，不纳入偏差判定，
    # 避免把作者粘贴带入的西文直接格式误报为"格式错误"。
    # ---------- 正文基准 ----------
    "Normal": {
        "jc": "both", "sz": 10.5,
        "eastAsia": "宋体",
    },
    "我的正文": {
        "jc": "both", "sz": 12, "line": 360, "lineRule": "auto", "firstLine": 420,
        "eastAsia": "宋体",
    },

    # ---------- 各级标题 ----------
    "heading 1": {   # 章标题：第X章
        "jc": "center", "sz": 15, "bold": True, "line": 360, "lineRule": "auto",
        "before": 480, "after": 360, "outlineLvl": 0, "pageBreakBefore": True,
        "eastAsia": "黑体",
    },
    "heading 2": {   # 节标题：1.1
        "jc": "left", "sz": 14, "bold": True, "line": 360, "lineRule": "auto",
        "before": 360, "after": 240, "left": 420, "outlineLvl": 1,
        "eastAsia": "黑体",
    },
    "heading 3": {   # 小节标题：1.1.1
        # 注：样式定义未显式设置 jc，继承 Normal 的两端对齐
        "sz": 12, "bold": True, "line": 360, "lineRule": "auto",
        "before": 240, "after": 240, "left": 147, "outlineLvl": 2,
        "eastAsia": "黑体",
    },
    "heading 4": {   # 四级标题：1.1.1.1
        "jc": "left", "sz": 12, "bold": True, "line": 360, "lineRule": "auto",
        "left": 147, "outlineLvl": 3,
        "eastAsia": "黑体",
    },

    # ---------- 篇级大标题（摘要/Abstract/目录/Contents/参考文献/致谢） ----------
    "我的标题": {
        "jc": "center", "sz": 18, "bold": True, "line": 720, "lineRule": "exact",
        "after": 480, "outlineLvl": 0, "pageBreakBefore": True,
        "eastAsia": "黑体",
    },
    "论文标题": {   # 封面"硕士学位论文"级别的大字
        "sz": 22, "eastAsia": "黑体",
    },
    "论文题目": {   # 封面题目
        "sz": 22, "eastAsia": "黑体",
    },

    # ---------- 图表 ----------
    "图标题": {
        "jc": "center", "sz": 12, "bold": True, "line": 360, "lineRule": "auto",
        "after": 320, "firstLine": 0,
    },
    "表标题": {
        "jc": "center", "sz": 12, "bold": True, "line": 360, "lineRule": "auto",
        "before": 320, "firstLine": 0,
    },
    "表格内文字": {
        "jc": "center", "sz": 10.5, "line": 240, "lineRule": "auto", "firstLine": 0,
    },
    "图表说明": {   # 图表下方"注：……"
        "sz": 7.5, "line": 0, "lineRule": "atLeast", "after": 120,
    },

    # ---------- 参考文献 ----------
    "文献列表": {   # 五号、单倍行距、无首行缩进，编号 [1] 由自动编号生成
        "sz": 10.5, "firstLine": 0,
    },

    # ---------- 页眉页脚 ----------
    "header": {"jc": "center", "sz": 10.5},
    "footer": {"sz": 9},
}

# ============================================================================
# 4. 段落角色规范
# ============================================================================
# 角色识别见 docx_kit.detect_role()；此处定义各角色应套用的样式名。
# styleId 会由 fmt_apply 按目标文档实际情况解析（同名样式保留其 ID）。
ROLE_TO_STYLE = {
    "cover_title":      "论文标题",      # 封面"硕士学位论文"字样
    "cover_subject":    "论文标题",      # 封面题目
    "major_title":      "我的标题",      # 摘要/Abstract/目录/参考文献/致谢
    "chapter_title":    "heading 1",
    "section_title":    "heading 2",
    "subsection_title": "heading 3",
    "subsubsection_title": "heading 4",
    "figure_caption":   "图标题",
    "table_caption":    "表标题",
    "table_note":       "图表说明",
    "table_body":       "表格内文字",
    "reference_item":   "文献列表",
    "body":             "我的正文",
    "abstract_cn":      "我的正文",
    "abstract_en":      "我的正文",
    "keywords":         "我的正文",
}

# 篇级大标题的文本（用于识别 major_title 角色）
MAJOR_TITLE_TEXTS = [
    "摘要", "摘 要", "摘  要",
    "Abstract",
    "目录", "目 录", "目  录",
    "Contents",
    "参考文献",
    "致谢",
    "附录",
]

# 需要重新起页码的分节序号（0-based，按文档中各 sectPr 出现顺序）
# 说明：样例文档 6 节 —— 0 封面 / 1 声明 / 2 摘要+目录（罗马）/ 3 正文（阿拉伯重启）/ 4 参考文献+致谢 / 5 文末
# 具体节数因文档而异，fmt_apply 按"是否含正文第一章"重新定位

# ============================================================================
# 5. 标题编号机制与冲突检测
# ============================================================================
# 规范（与样例文档一致）：
#   章标题 —— 样式 heading 1 **不带**编号；编号来自**段落级** numPr(numId, ilvl=0)，
#             生成「第X章」。故段落文本应为纯标题（如「绪论」）。
#   节及以下 —— 样式 heading 2/3/4 **自带**样式级 numPr(ilvl=1/2/3)，
#             自动生成「1.1」「1.1.1」。故段落文本同样应为纯标题。
#
# 若标题文本里仍保留手写编号，套用规范样式后 Word 会渲染成
# 「1.1 1.1 选题背景」这类重复编号（实测复现过），必须检测出来。
CONFLICT_HEADING_ROLES = (
    "section_title", "subsection_title", "subsubsection_title",
)

RE_MANUAL_CHAPTER = re.compile(r"^\s*第\s*[一二三四五六七八九十百零〇0-9]+\s*[章篇]")
RE_MANUAL_DECIMAL = re.compile(r"^\s*\d+(?:\.\d+)*[\s、:：.．]")
RE_MANUAL_OUTLINE = re.compile(r"^\s*\d{1,2}[ \u3000]+\S")

# 「合并标题」：一段里挤了两条标题（如「金融监管的必要性 银行业监管的必要性」），
# 常见于作者把两条提纲并到一行、或删改标题时漏删。
# 判据收得很紧，避免误报：仅限章/节级；以空白分隔；两侧各至少 5 个字符且都不含空白；
# 这样「摘  要」「第三章 研究方法」这类正常标题不会被误判。
RE_MERGED_HEADING = re.compile(
    r"^([\u4e00-\u9fffA-Za-z0-9（）()]{5,})[\s\u3000]+([\u4e00-\u9fffA-Za-z0-9（）()]{5,})$")
