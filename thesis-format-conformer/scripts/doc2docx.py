# -*- coding: utf-8 -*-
"""
把 Word 97-2003 的 .doc 另存为 .docx（OOXML），供 fmt_check / fmt_apply 处理。

用法：
    python doc2docx.py <输入.doc> [输出.docx]

依赖：Windows + 已安装 Microsoft Word + pywin32（`pip install pywin32`）。
     Word 必须存在 —— 这一步是「让 Word 自己另存」，不是第三方格式解析，
     因此能无损保留分节、页眉页脚、表格、域。

转换后常见现象（属正常，不是转换出错）：
    - 所有段落都"没有命名样式"、只有直接格式 —— 这是 .doc 的常态，
      也正是 fmt_apply 要按角色重新套样式的原因；
    - 章标题段首可能带手动分页符（w:br type="page"）。
"""
import argparse
import os
import sys


def convert(src, dst, overwrite=False):
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if not os.path.exists(src):
        raise SystemExit("输入文件不存在: %s" % src)
    if os.path.abspath(src) == dst:
        raise SystemExit("输入与输出不能是同一个文件")
    if os.path.exists(dst) and not overwrite:
        raise SystemExit("输出已存在，拒绝覆盖（加 --overwrite）: %s" % dst)

    try:
        import win32com.client
    except ImportError:
        raise SystemExit(
            "缺少 pywin32。请先安装：\n"
            "    pip install pywin32\n"
            "（本转换依赖本机已安装的 Microsoft Word，非 Windows 环境请改用 LibreOffice）")

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
        try:
            # wdFormatDocumentDefault = 16 -> .docx
            doc.SaveAs2(dst, FileFormat=16)
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()
    return dst


def main():
    ap = argparse.ArgumentParser(description=".doc 转 .docx（调用本机 Word 另存）")
    ap.add_argument("doc", help="输入 .doc")
    ap.add_argument("out", nargs="?", help="输出 .docx（默认与输入同目录同名）")
    ap.add_argument("--overwrite", action="store_true", help="允许覆盖已存在的输出")
    args = ap.parse_args()

    out = args.out or os.path.splitext(args.doc)[0] + ".docx"
    dst = convert(args.doc, out, overwrite=args.overwrite)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print("已转换：%s（%d 字节）" % (dst, os.path.getsize(dst)))


if __name__ == "__main__":
    main()
