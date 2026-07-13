# pitch-review — 推票周度复盘器（自包含 · 可分发）

把"上周按短/中/长期推的票"逐一验证，产出**图文并茂的复盘报告**（写进 Obsidian 或任意 Markdown）：开篇导语 → 大盘复盘 → 板块复盘 → 个股复盘（可选加下周推荐）。

## 前置（很轻）

- **Python 3.8+**，**无需任何第三方包**——`scripts/` 全部只用标准库。
- 一个能跑 Agent + 读写 Markdown 的环境（Claude Code / 兼容宿主）。
- 图表用 Obsidian 的 `echarts` 代码块渲染；其它 Markdown 环境需支持 ECharts 代码块或自行换渲染器。
- 可选：需走代理时设标准 `HTTPS_PROXY` 环境变量，脚本自动识别（不在文件里硬编码任何地址）。

## 5 分钟跑通（开箱即用，零外部账号/密钥）

数据默认走 **Yahoo Finance**（免 API key、免注册，`scripts/fetch_bars.py` 纯标准库实现）：

```bash
# 1) 取行情（Yahoo 代码：指数带 ^，港股 .HK，台股 .TW，韩股 .KS）
python scripts/fetch_bars.py --interval 1d --range 1mo --spec ^GSPC ^SOX 2330.TW > spec.json

# 2) 编辑 spec.json：给每只补 title/sub/unit；复盘图设 "highlight_start":"MM/DD" 且各 series "highlight":true

# 3) 生成 ECharts 蜡烛图块
python scripts/kcharts.py spec.json > blocks.md      # 把 blocks.md 里的 ```echarts 块插进报告

# 4) 交付前校验（所有图 JSON 合法；日K 高亮标签落在 X 轴内）
python scripts/validate_echarts.py 你的报告.md
```

其余四类图（跨市场 bar / rebased 折线 / 行业全景 bar / 个股跌幅 bar）按 `references/data-and-charts.md` 的模板手写 JSON，同样用 `validate_echarts.py` 校验。

先离线试脚本（不联网）：`python scripts/kcharts.py references/example/spec.example.json`。

## 数据源：一个默认，可按契约替换

本 skill 只依赖一份「数据层契约」（见 `references/data-and-charts.md` §一.5），默认实现即自带脚本：

| 角色 | 契约 | 默认实现（无 token） | 想换就换成 |
|---|---|---|---|
| 行情 / 日K | `symbol, interval, count → bars[{t,o,h,l,c,v}]` | `scripts/fetch_bars.py`（Yahoo） | yfinance / akshare / 券商 API 等 |
| 财报三表 / 一致预期 | `ticker → 三表 + 现价/PE + 一致预期` | SEC EDGAR / 免费源 / 手填 | 任意结构化源 |
| A股申万一级行业 | 31 行业区间涨跌（同口径） | 免费源（如 akshare，需 `pip install akshare`） | 任意结构化源 |
| 板块催化 / 新闻 | 关键词检索 | 宿主的 Web 搜索 | 任意搜索源 |

> Yahoo 覆盖美/港/台/韩个股、主要指数、GICS 行业 ETF；**A股申万行业指数 Yahoo 没有**，用免费源或你自己的结构化源补（契约见 references）。

## 无密钥承诺

本 skill 目录内**任何文件都不含 token / 密钥 / 私有服务器地址**。密钥只应在你本机的临时环境变量里出现，绝不写进 skill、报告或会同步的目录。

## 目录

```
pitch-review/
├─ SKILL.md                     工作流（Step 0–6 + 可选 Step 5B 下周推荐），每步带规则
├─ README.md                    本文件（分发/上手/契约）
├─ scripts/
│  ├─ fetch_bars.py             零依赖默认行情源（Yahoo，免 key）
│  ├─ kcharts.py                日K/周K 生成器（红涨绿跌，可选复盘高亮）
│  └─ validate_echarts.py       交付前 ECharts 校验
└─ references/
   ├─ data-and-charts.md        取数配方 + 5 类图模板 + 写作红线 + 数据层契约
   └─ example/spec.example.json 离线样例（无需联网即可跑 kcharts）
```

## 边界

只做**复盘**（验证已推的票，并可把已定的下周清单写成推荐段）。不做全市场选股、不做深度覆盖/估值建模——本 skill 只吃清单、客观记录兑现结果。
