# -*- coding: utf-8 -*-
"""
自包含默认行情源：Yahoo Finance chart 接口（免 API key、纯标准库、零第三方包）。
满足 references「数据层契约」的行情源角色，让拿到本 skill 的任何人开箱即用，
不依赖任何私有服务器、账号或 token。

输出即 kcharts.py 能吃的格式：
  --raw  : {symbol: [{t,o,h,l,c,v}, ...]}                （自己再组 spec）
  --spec : {highlight_start, series:[{title,unit,interval,bars}]}  （可直接喂 kcharts.py）

用法：
  python fetch_bars.py --interval 1d --range 1mo --spec ^GSPC ^SOX 2330.TW > spec.json
  python fetch_bars.py --interval 1wk --range 3y --spec NVDA TSM > spec.json
  python fetch_bars.py --raw GLW MU 0700.HK > bars.json
  # 可选：需走代理时设标准环境变量 HTTPS_PROXY（不在脚本内硬编码任何地址）

Yahoo 代码速查（复盘常用）：
  指数  上证 000001.SS｜深成 399001.SZ｜创业板 399006.SZ｜沪深300 000300.SS｜科创50 000688.SS
        恒生 ^HSI｜恒生科技 ^HSTECH｜国企 ^HSCE｜道指 ^DJI｜标普 ^GSPC｜纳指 ^IXIC｜费半 ^SOX
  参考  VIX ^VIX｜美债10Y ^TNX｜美元 DX-Y.NYB｜黄金 GC=F｜原油 CL=F
  行业  GICS ETF 直接用 XLK/XLF/XLE/XLV/XLI/XLC/XLY/XLP/XLU/XLB/XLRE
  个股  美股直接代码(GLW/MU/TSM…)｜港股 0700.HK｜台股 2330.TW｜韩股 000660.KS/005930.KS
  A股申万一级 31 行业区间涨跌不在 Yahoo；用 references 里的结构化源（akshare 免费可替换）。
"""
import sys, json, argparse, urllib.request, urllib.parse, os

BASE = "https://query1.finance.yahoo.com/v8/finance/chart/"

def _opener():
    proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
    if proxy:
        return urllib.request.build_opener(urllib.request.ProxyHandler({"http": proxy, "https": proxy}))
    return urllib.request.build_opener()

def fetch(symbol, interval, rng):
    url = f"{BASE}{urllib.parse.quote(symbol)}?range={rng}&interval={interval}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    raw = _opener().open(req, timeout=30).read()
    r = json.loads(raw)["chart"]["result"][0]
    ts = r["timestamp"]
    q = r["indicators"]["quote"][0]
    bars = []
    for i, t in enumerate(ts):
        o, h, l, c = q["open"][i], q["high"][i], q["low"][i], q["close"][i]
        if None in (o, h, l, c):
            continue
        v = (q.get("volume") or [None] * len(ts))[i] or 0
        bars.append({"t": int(t), "o": round(o, 4), "h": round(h, 4), "l": round(l, 4), "c": round(c, 4), "v": v})
    return bars

def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="+", help="Yahoo 代码，如 ^GSPC ^SOX 2330.TW")
    ap.add_argument("--interval", default="1d", help="1d / 1wk / 1mo")
    ap.add_argument("--range", dest="rng", default="1mo", help="1mo / 6mo / 1y / 3y / 5y")
    ap.add_argument("--spec", action="store_true", help="输出 kcharts spec 结构")
    ap.add_argument("--raw", action="store_true", help="输出 raw bars 字典（默认；与 --spec 二选一）")
    a = ap.parse_args()

    out = {}
    errs = {}
    for s in a.symbols:
        try:
            out[s] = fetch(s, a.interval, a.rng)
        except Exception as e:
            errs[s] = repr(e)[:150]
    for s, e in errs.items():
        print(f"[warn] {s}: {e}", file=sys.stderr)

    if a.spec:
        iv = "1W" if a.interval.startswith("1w") else ("1M" if a.interval == "1mo" else "1D")
        spec = {"highlight_start": "", "series": [
            {"title": s, "sub": "", "unit": "", "interval": iv, "highlight": False, "bars": b}
            for s, b in out.items() if b
        ]}
        print(json.dumps(spec, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(out, ensure_ascii=False))

if __name__ == "__main__":
    main()
