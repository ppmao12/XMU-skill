# thesis-format-conformer

> 把中文学位论文的 `.docx` 格式，一键对齐 **厦门大学** 规范（也支持换到其他学校）。

这是 `XMU-skill` 的核心实现。安装方式与完整说明见 [仓库主页 README](../README.md)；本文件只讲**这个 skill 本身怎么用**。

---

## 30 秒上手

```bash
pip install lxml pillow

# 1. 体检（只读，不改文件）
python scripts/fmt_check.py "论文.docx" --json 报告.json

# 2. 修复
python scripts/fmt_apply.py "论文.docx" -o "论文-已规范化.docx" --report 变更日志.txt

# 3. 分节 + 页码 + 页眉页脚
python scripts/fmt_hf.py "论文-已规范化.docx" -o "论文-最终.docx" --title "论文题目"

# 4.（可选）换封面校徽
python scripts/fmt_cover.py "论文-最终.docx" -o "论文-交付.docx" --logo assets/xiamen-university-logo.png

# 5. 复核 —— 第六节编号冲突数须为 0
python scripts/fmt_check.py "论文-交付.docx"
```

**顺序不能乱**：`fmt_check` → `fmt_apply` → `fmt_hf` → `fmt_cover` → `fmt_check`。
`fmt_hf` 依赖 `fmt_apply` 已套好的样式来定位章标题。

**交付前**：在 Word 里 `Ctrl+A` → `F9` 更新域（目录/页码/页眉章名都是域，不刷新显示旧值）。

---

## 文件结构

```
thesis-format-conformer/
├── SKILL.md                    # 给 agent 读的说明书（触发词、工作流、技术立场）
├── README.md                   # 本文件（给人读的速查）
├── assets/
│   ├── reference-styles.xml    # 规范样式库（90 KB，样式同步的来源）
│   ├── reference-numbering.xml # 规范编号库（多级标题 + 参考文献编号）
│   └── xiamen-university-logo.png
├── references/
│   └── format-spec.md          # 规范全文（实测数据 + 已知局限）
└── scripts/
    ├── fmt_check.py            # 只读诊断，六节报告
    ├── fmt_apply.py            # 执行修复：页面 / 样式 / 段落 / 编号
    ├── fmt_hf.py               # 分节 + 页码 + 页眉页脚
    ├── fmt_cover.py            # 换封面校徽 / 校名图
    ├── doc2docx.py             # .doc → .docx（调本机 Word，无损）
    ├── spec_rules.py           # 规范常量 —— 改规范只改这里
    └── docx_kit.py             # OOXML 工具库（样式链 / 角色识别 / 分节解析）
```

---

## 依赖

| 包 | 用途 | 是否必需 |
|---|---|---|
| `lxml` | 解析 OOXML | **必需** |
| `pillow` | 换封面校徽（`fmt_cover.py`） | 可选 |
| `pywin32` | `.doc → .docx`（`doc2docx.py`，另需本机装 Microsoft Word） | 可选 |

---

## 规范速查（厦门大学）

| 项目 | 值 |
|---|---|
| 纸张 / 边距 | A4；上 3.0 / 下 2.54 / 左 3.17 / 右 3.17 cm |
| 正文 | 宋体 · 小四 12pt · 1.5 倍行距 · 首行缩进 2 字符 |
| 章标题 | 黑体 · 小三 15pt · 加粗 · 居中 · 段前分页 |
| 节标题 | 黑体 · 四号 14pt · 加粗 · 左对齐 |
| 小节标题 | 黑体 · 小四 12pt · 加粗 |
| 篇级大标题 | 黑体 · 小二 18pt · 加粗 · 居中 |
| 表格内文字 | 宋体 · 五号 10.5pt · 居中 · 单倍行距 |
| 参考文献 | 宋体 · 五号 · 单倍行距 · 无首行缩进 · 编号 `[1]` 自动生成 |
| 页码 | 封面无 → 前置大写罗马 → 正文阿拉伯重新起 1 |

完整规范（含每条数值的来源文档）见 [`references/format-spec.md`](references/format-spec.md)。

---

## 换到其他学校

规范数值集中在两处：`scripts/spec_rules.py` 与 `assets/reference-*.xml`。

1. 用新模板的 `styles.xml` / `numbering.xml` 替换 `assets/` 下同名文件；
2. 改 `spec_rules.py` 的 `PAGE_*` / `MARGIN_*` / `STYLE_SPECS` / `ROLE_TO_STYLE`；
3. 拿该模板跑一次 `fmt_check.py`，确认自检问题趋近于零再投入使用。

**只换学校标识**（规范不变）改三个常量即可，或运行时传参：

```bash
python scripts/fmt_hf.py ... --school-header "XX大学本科生毕业论文"
python scripts/fmt_cover.py ... --logo "assets/你的校徽.png"
```

---

## 几条会踩的坑

1. **标题里的手写编号必须剥离** —— 节及以下的编号挂在**样式**上，只删段落 `numPr` 挡不住，结果是「1.1 1.1 标题」。`fmt_apply` 自动处理。
2. **章标题必须写 `w:ind left="0"`** —— 编号定义自带 8.5 cm 左缩进，不抵消的话长标题当场折行。`fmt_check` 查不出来，只能靠 Word 渲染看。
3. **奇偶页页脚必须 `default` 与 `even` 同时挂** —— 否则偶数页没有页码。
4. **域结果有缓存** —— 改完务必 `F9`。
5. **封面识别靠启发式** —— 看 `fmt_apply` 日志有无「封面/声明区保护」行，缺了或范围异常就要人工核对。

其余坑与人工介入项见 [`references/format-spec.md`](references/format-spec.md) 第 10 节。

---

## License

[MIT](../LICENSE)
