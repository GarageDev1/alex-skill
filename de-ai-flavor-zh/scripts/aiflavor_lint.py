#!/usr/bin/env python3
"""AI 味候选标记器 —— de-ai-flavor-zh skill 的辅助工具。

只拎"高频可疑形态"当线索,不下判决。去 AI 味高度依赖语境,
一个句式是不是真有问题,最终按 SKILL.md 的规则人工判。
所以本脚本始终退出 0(advisory),不阻断任何流程。

诚实声明:正则只抓得住有固定形态的那几类(反转金句、设问开头、
比喻二选一……),抓不住"架子句""贴标签"这种要读懂意思才看得出的 AI 味。
那一类靠 SKILL.md 的三个诊断问题人工过。

用法:
    python aiflavor_lint.py <文件/目录> [<更多> ...]
    传目录则递归扫描 Markdown、Python、JSON、YAML、HTML 和 TXT 文件。

退出码:始终 0(只提示,不阻断)。
"""

import os
import re
import sys

# 每条:(类型, 正则, 提示)。逐行扫;标 ^ 的按行首匹配。
PATTERNS = [
    ("反转金句", re.compile(r"这不是.{0,18}[，,].{0,12}是"),
     "“这不是 X，是 Y”反转金句——删掉，直接讲 Y 是什么 + 机制"),
    ("为转而转", re.compile(r"不是.{1,15}[，,]?\s*而是"),
     "“不是 X，而是 Y”——X 若是没人真信的稻草人就拉平，只点名真实立场才留"),
    ("设问开头", re.compile(r"^\s*那.{0,24}[?？]")),
    ("开场垫话", re.compile(r"^\s*(答案很|答案是|问题来了|值得一问|说到底|归根结底|本质上|说白了)")),
    ("设问自答", re.compile(r".{0,30}[?？]\s*$")),  # 行尾问号:疑似设问,自答紧随
    ("比喻二选一", re.compile(r"是.{2,20}还是.{2,20}[。?？]"),
     "“是 A 还是 B”若 A、B 都是比喻——换成两种字面情况 + 怎么区分"),
    ("破折号抖包袱", re.compile(r"——")),
    ("否定对仗", re.compile(r"不仅仅是.{0,20}更是|与其说是.{0,20}不如说是")),
    ("生造组合词", re.compile(r"接法|软依赖|交付门|物理锚|锚链|多镜头|三角验证|隐含价阶梯|历史底座|镜头体检|情景翻档|翻档|链断"),
     "生造组合词或内部工程黑话——拆成对象、动作、条件和结果"),
    ("内部工程黑话", re.compile(r"喂给|挂(?:上|到|在)|公司切片|业务盘子|吃哪一刀|长进现价"),
     "内部工程黑话——改成具体的引用、驱动、份额或计算关系"),
]

DEFAULT_HINT = {
    "设问开头": "设问自答开头——删设问，直接上结论",
    "开场垫话": "开场垫话——删掉，直接给结论",
    "设问自答": "行尾问号，疑似设问自答——若后面紧跟自答，删问句直接给答案",
    "破折号抖包袱": "破折号——少用“停顿—揭晓”，多数能换成逗号或句号直接陈述",
    "否定对仗": "否定式排比拔高——拉平成正面陈述",
}

SCAN_EXTS = (".md", ".txt", ".py", ".json", ".yaml", ".yml", ".html")


def scan_file(path):
    try:
        with open(path, encoding="utf-8", errors="replace") as fh:
            lines = fh.readlines()
    except OSError as e:
        print(f"[warn] 读不了 {path}: {e}", file=sys.stderr)
        return []

    hits = []
    for i, line in enumerate(lines, 1):
        s = line.rstrip("\n")
        if not s.strip():
            continue
        for entry in PATTERNS:
            kind, pat = entry[0], entry[1]
            hint = entry[2] if len(entry) > 2 else DEFAULT_HINT.get(kind, "")
            if pat.search(s):
                hits.append((i, kind, hint, s.strip()))
    return hits


def iter_files(paths):
    for p in paths:
        if os.path.isdir(p):
            for root, _, files in os.walk(p):
                for f in files:
                    if f.endswith(SCAN_EXTS):
                        yield os.path.join(root, f)
        elif os.path.isfile(p):
            yield p
        else:
            print(f"[warn] 路径不存在,跳过:{p}", file=sys.stderr)


def main(argv):
    paths = argv[1:]
    if not paths:
        print(__doc__)
        return 0

    total = 0
    for path in iter_files(paths):
        hits = scan_file(path)
        if not hits:
            continue
        total += len(hits)
        print(f"\n● {path}")
        for lineno, kind, hint, snippet in hits:
            print(f"  [{kind}] L{lineno}  {hint}")
            print(f"      > {snippet}")

    print("\n" + "=" * 60)
    if total:
        print(f"标出 {total} 处候选形态。这些只是线索,不是判决——")
        print("逐句按 SKILL.md 的三个诊断问题和规则人工判:删掉空句式后,")
        print("信息有没有损失?没损失就拉平。数字和事实一个不动。")
    else:
        print("没拎到固定形态的候选。注意:架子句、贴标签这类靠读意思才看得出的 AI 味,")
        print("正则抓不住,仍要按 SKILL.md 人工过一遍。")
    return 0  # 始终不阻断


if __name__ == "__main__":
    sys.exit(main(sys.argv))
