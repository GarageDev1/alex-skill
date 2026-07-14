#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
parse_blocks 回归测试 —— 锁定"标题不丢"。

跑法:
  python tests/test_parse_blocks.py

每一行是一个 (用例名, markdown, 期望被识别出的标题列表)。
期望 [(level, text), ...] 是文档里所有应被解析为 heading 的标题,顺序即出现顺序。
"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
sys.stdout.reconfigure(encoding="utf-8")

import md_to_docx as M


def headings(md):
    """返回解析出的所有标题 [(level, text)]"""
    return [
        (b["level"], b["text"])
        for b in M.parse_blocks(md)
        if b["type"] == "heading"
    ]


def block_types(md):
    return [b["type"] for b in M.parse_blocks(md)]


CASES = [
    # ---- ATX 标题 ----
    ("标准 H2",            "## 标题",                [(2, "标题")]),
    ("H1~H6 各级",        "# H1\n## H2\n### H3",   [(1, "H1"), (2, "H2"), (3, "H3")]),
    ("前导 1 空格",        " ## 标题",               [(2, "标题")]),
    ("前导 2 空格",        "  ## 标题",              [(2, "标题")]),
    ("前导 3 空格(上限)",  "   ## 标题",             [(2, "标题")]),
    ("前导 tab",           "\t## 标题",              [(2, "标题")]),
    ("尾部空格",           "## 标题   ",             [(2, "标题")]),
    ("标题含竖线",         "## A | B",               [(2, "A | B")]),
    ("标题含 inline 粗体", "## **粗** 标题",         [(2, "**粗** 标题")]),

    # ---- Setext 标题 ----
    ("Setext H1 (===)",   "标题一\n===\n正文",      [(1, "标题一")]),
    ("Setext H2 (---)",   "标题二\n---\n正文",      [(2, "标题二")]),
    ("Setext 多个=号",     "标题\n=========",        [(1, "标题")]),

    # ---- 标题紧跟其他块(都不该被吞) ----
    ("表格后标题",         "| a |\n| --- |\n| 1 |\n## T",  [(2, "T")]),
    ("列表后标题",         "- 项目\n## T",                  [(2, "T")]),
    ("段落后标题",         "正文段。\n## T",                [(2, "T")]),
    ("引用后标题",         "> 引用\n## T",                  [(2, "T")]),
    ("两段之间标题",       "段一\n\n## T\n\n段二",          [(2, "T")]),

    # ---- 代码块 ----
    ("闭合代码块后标题",   "```\ncode\n```\n## T",   [(2, "T")]),

    # ---- 不应误判为标题(负例) ----
    ("4空格不判为标题",    "    ## 不是标题",        []),  # CommonMark: 4空格是缩进代码块
    ("#无空格不判为标题",  "##无空格",               []),  # CommonMark 非法 ATX
]

# ---- 反向断言:某些结构不应被当成标题 ----
NEG_CASES = [
    # 单独 --- 是 hr,不是 Setext(前一行不是段落文本)
    ("单独---是hr", "---", ["hr"]),
    ("单独===不成标题", "===", None),  # === 单独出现:不算标题即可(段落或忽略都行)
]


def run():
    fail = 0
    for name, md, expected in CASES:
        got = headings(md)
        ok = got == expected
        flag = "OK  " if ok else "FAIL"
        if not ok:
            fail += 1
        print(f"{flag} {name:24s} got={got}")
    for name, md, expected_types in NEG_CASES:
        types = block_types(md)
        has_heading = "heading" in types
        if has_heading:
            fail += 1
            print(f"FAIL {name:24s} 误判为 heading, types={types}")
        else:
            print(f"OK   {name:24s} 未误判 heading, types={types}")

    print(f"\n{'✅ ALL PASS' if fail == 0 else f'❌ {fail} FAILED'}")
    return fail


if __name__ == "__main__":
    sys.exit(1 if run() else 0)
