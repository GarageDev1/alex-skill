# -*- coding: utf-8 -*-
"""
交付前验收：扫描一个 .md 里所有 ```echarts 块。
- 硬门槛（失败即 exit 1）：每个块的 JSON 必须能 json.loads。
- 软提示（只打印、不拦）：带 markArea 的蜡烛图，检查高亮起止标签是否落在 xAxis 里
  （复盘日K 应有高亮；下周推荐类周K 没有 markArea 属正常，不提示）。
用法：python validate_echarts.py <md路径>
"""
import sys, re, json

def main():
    t = open(sys.argv[1], encoding="utf-8").read()
    blocks = re.findall(r"```echarts\n(.*?)\n```", t, re.S)
    json_fail = 0
    ma_warn = 0
    k = 0
    for i, b in enumerate(blocks, 1):
        try:
            o = json.loads(b)
        except Exception as e:
            print(f"  block {i}: JSON FAIL -> {e}"); json_fail += 1; continue
        series = o.get("series", [{}])
        s0 = series[0] if series else {}
        if s0.get("type") == "candlestick":
            k += 1
            if "markArea" in s0:  # 只校验声称要高亮的蜡烛图
                dates = o.get("xAxis", {}).get("data", [])
                ma = s0["markArea"].get("data", [[{}, {}]])[0]
                st, en = ma[0].get("xAxis"), ma[1].get("xAxis")
                if not ((st in dates) and (en in dates)):
                    print(f"  [warn] Kchart[{i}] markArea 标签不在 X 轴内: start={st} end={en}")
                    ma_warn += 1
    print(f"echarts blocks: {len(blocks)} | candlesticks: {k} | JSON failures: {json_fail} | markArea warnings: {ma_warn}")
    sys.exit(1 if json_fail else 0)

if __name__ == "__main__":
    main()
