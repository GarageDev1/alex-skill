#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
_compute_col_widths 单元测试 —— 锁住"列宽按内容比例分配"。

跑法:
  python tests/test_tables.py

核心不变量:
  1. 每列宽度 > 0
  2. 列宽总和 = 正文总宽(归一化)
  3. 单列表 = 整个正文宽
  4. 内容长的列 > 内容短的列(这是用户要的核心:不再等宽)
  5. 列数很多时,窄列不被强行拉到 min 导致整体失去区分
"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
sys.stdout.reconfigure(encoding="utf-8")

import md_to_docx as M

TOTAL = 6.0


def approx(a, b, eps=0.02):
    return abs(a - b) < eps


def run():
    fail = 0

    def check(name, cond, detail=""):
        nonlocal fail
        flag = "OK  " if cond else "FAIL"
        if not cond:
            fail += 1
        print(f"{flag} {name} {detail}")

    # 1. 基本不变量:正数 + 总和=TOTAL
    w = M._compute_col_widths(["a", "b", "c"], [["1", "2", "3"]], TOTAL)
    check("列宽全为正", all(x > 0 for x in w), f"w={[round(x,2) for x in w]}")
    check("总和=TOTAL", approx(sum(w), TOTAL), f"sum={round(sum(w),3)}")

    # 2. 单列 = 整宽
    w = M._compute_col_widths(["only"], [["x"]], TOTAL)
    check("单列表=整宽", approx(w[0], TOTAL), f"w={w}")

    # 3. 核心:内容长的列明显宽于内容短的列
    header = ["短", "这是一个内容很长的列标题文字特别多"]
    data = [["x", "yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy"]]
    w = M._compute_col_widths(header, data, TOTAL)
    check("长列 > 短列", w[1] > w[0] * 2, f"短={round(w[0],2)} 长={round(w[1],2)}")

    # 4. 等长内容应接近等宽(不误伤)
    w = M._compute_col_widths(["aa", "bb", "cc"], [["11", "22", "33"], ["44", "55", "66"]], TOTAL)
    spread = max(w) - min(w)
    check("等长内容接近等宽", spread < 0.5, f"spread={round(spread,2)} w={[round(x,2) for x in w]}")

    # 5. 列数很多(10列,内容长短悬殊)时仍有区分度,不是全被 clamp 成等宽
    header = ["方法ID", "层级", "状态", "方法名", "代表样本", "核心证据", "适用场景", "风险", "维度", "迁移"]
    data = [[
        "M01", "战略层", "核心", "这是一个较长的方法名称",
        "7638902116808461595（创作者vs创业者）",
        '"创作者看量，创业者看率"', "任何知识类IP初期定位",
        "需要真的有创业/商业经历支撑", "IP定位", "可直接复用"
    ]]
    w = M._compute_col_widths(header, data, TOTAL)
    check("10列密集表最长/最短 > 2.5x", max(w) / min(w) > 2.5,
          f"最长={round(max(w),2)} 最短={round(min(w),2)} 比={round(max(w)/min(w),2)}")
    check("10列代表样本列(视频ID)是最宽之一", w[4] >= sorted(w)[-2],
          f"代表样本={round(w[4],2)}")

    print(f"\n{'✅ ALL PASS' if fail == 0 else f'❌ {fail} FAILED'}")
    return fail


if __name__ == "__main__":
    sys.exit(1 if run() else 0)
