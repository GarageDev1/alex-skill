# -*- coding: utf-8 -*-
"""
可复用 K 线生成器：把 OHLCV 数据转成 ECharts 蜡烛图块（中式红涨绿跌）。
数据驱动、零硬编码、仅用标准库——换 spec.json 即可重跑，可用于任何标的/区间/周期。

支持两种用途：
  A) 复盘（Part2/Part4）：日K + "复盘区间"黄色高亮（highlight_start 落在 X 轴内）。
  B) 下周推荐（Part5）：按持仓期用不同周期，通常不高亮——
       短线=日K约1月(interval 1D, count≈25)、中线=周K约6月(1W, count≈30)、
       长线=周K约3年(1W, count≈160)。

用法：
  1) 用行情源取每只的 bars（字段 t/o/h/l/c，t=Unix 秒 UTC；来源不限，见 references 数据层契约）。
  2) 写 spec.json：
     {
       "highlight_start": "06/29",        # 可选，复盘区间起点(MM/DD)；推荐段可省
       "series": [
         {"title": "台积电 TSM 周K·近三年",
          "sub": "三年 +315%…",
          "unit": "$",
          "interval": "1W",              # 可选，"1D"(默认)/"1W"；决定日期格式与标签密度
          "highlight": false,            # 可选，默认 true(有 highlight_start 才画高亮)
          "bars": [{"t":1687267800,"o":104.18,"h":104.88,"l":101.0,"c":101.91}, ...]}
       ]
     }
  3) python kcharts.py spec.json  > blocks.md    # 打印所有 ```echarts 块
调用方(agent)再把块插到 md 对应位置，最后必过 validate_echarts.py。
"""
import sys, json
from datetime import datetime, timezone

try:                                      # Windows 终端默认 GBK，重定向到文件会污染中文；强制 UTF-8
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

UP, DOWN = "#ee6666", "#3ba272"  # 红涨 / 绿跌（中式）


def _fmt(t, mode):
    dt = datetime.fromtimestamp(t, tz=timezone.utc)
    return dt.strftime("%y/%m" if mode == "ym" else "%m/%d")


def _axis_style(interval, n):
    """按周期与 bar 数决定日期格式与标签密度，避免多年周K标签糊成一团。"""
    if interval == "1W" and n > 60:      # 多年周K：YY/MM，约每季一个标签
        return "ym", max(1, n // 12)
    if interval == "1W":                 # 数月周K：MM/DD，稀一点
        return "md", 3
    return "md", 2                        # 日K


def block(s, default_start):
    bars = s["bars"]
    interval = s.get("interval", "1D")
    mode, label_interval = _axis_style(interval, len(bars))
    dates = [_fmt(b["t"], mode) for b in bars]
    kdata = [[round(b["o"], 4), round(b["c"], 4), round(b["l"], 4), round(b["h"], 4)] for b in bars]
    opt = {
        "height": 360,
        "title": {"text": s["title"], "subtext": s.get("sub", ""), "left": "center",
                  "textStyle": {"fontSize": 14}, "subtextStyle": {"fontSize": 11, "color": "#666"}},
        "tooltip": {"trigger": "axis", "axisPointer": {"type": "cross"}},
        "grid": {"top": 62, "bottom": 40, "left": "11%", "right": "5%", "containLabel": True},
        "xAxis": {"type": "category", "data": dates, "boundaryGap": True,
                  "axisLabel": {"interval": label_interval}},
        "yAxis": {"scale": True, "name": s.get("unit", "")},
        "series": [{
            "type": "candlestick", "data": kdata,
            "itemStyle": {"color": UP, "color0": DOWN, "borderColor": UP, "borderColor0": DOWN},
        }],
    }
    # 复盘高亮（推荐段默认关；highlight=false 或无 start 则不画）
    start = s.get("start", default_start)
    if s.get("highlight", True) and start and start in dates:
        opt["series"][0]["markArea"] = {
            "silent": True,
            "itemStyle": {"color": "rgba(250, 200, 88, 0.20)"},
            "label": {"show": True, "position": "insideTop", "formatter": "复盘区间",
                      "color": "#b8860b", "fontSize": 11},
            "data": [[{"xAxis": start}, {"xAxis": dates[-1]}]],
        }
    json.loads(json.dumps(opt))  # 自检
    return "```echarts\n" + json.dumps(opt, ensure_ascii=False) + "\n```"


def main():
    spec = json.load(open(sys.argv[1], encoding="utf-8"))
    default_start = spec.get("highlight_start", "")
    print("\n\n".join(block(s, default_start) for s in spec["series"]))


if __name__ == "__main__":
    main()
