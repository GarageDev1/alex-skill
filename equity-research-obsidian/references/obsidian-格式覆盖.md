# Obsidian 格式覆盖(ICA 的 DOCX 格式 → Obsidian md)

ICA 是 DOCX 产物——字体字号、页边距、页眉页脚、封面双页、`w:rFonts`/`tblBorders`/numbering.xml、页码域、PDF 视觉核验(见 ICA `references/Word文档格式标准.md`)。本 skill 是 **Obsidian md**,这些 DOCX 专属格式**全部不适用**。本文件声明 Obsidian 的覆盖格式。

## 真相源分工

- **ICA `references/Word文档格式标准.md`** — **不加载**(DOCX 专属,Obsidian md 无字体/字号/页边距/分页)
- **ICA `assets/报告模板与版式.md`** — 中式四章骨架、封面信息项、体量基准**沿用**;但"封面双页(DOCX 分页)"改成 Obsidian 开头门面区块(见 §二)
- **本文件** — Obsidian md 格式(frontmatter / wiki link / 不分页 / 黑白名单 / 门面首页适配)

---

## 一、Obsidian md 基本格式

- **文件最前写 YAML frontmatter**(11 项,见 §三)。
- **跨公司提及用 `[[wiki link]]` 双链**(如 `[[NVIDIA]]`、`[[ASML|ASML 卡 EUV]]`),每篇 ≥3 处。
- **标题用 `#`/`##`/`###`,全中文,判断式**(ICA 研报写作规范.md)。
- **表格用纯 Markdown `|`**(不用 HTML/Word 表格)。
- **不分页**:md 无页概念,不写 page break / page number。ICA 的"25-45 页"在 md 里换成"1.5-3 万字 + 40-70 ECharts"(§四)。
- **不写 DOCX 指令**:无字体/字号/行距/页边距/页眉页脚/边框样式。
- **callout 禁用**(见 §五黑名单)。

---

## 二、门面首页(Obsidian 适配 ICA 封面双页)

ICA 封面分两页(DOCX 自然流式分页):第 1 页(主副标题/评级表/摘要)、第 2 页(盈利预测表/基本状况/股价图)。Obsidian md **不分页**,把封面信息收成**开头门面区块**(紧接 frontmatter、正文之前):

```markdown
# {公司}首次覆盖报告（{ticker}）
# {一句话 equity story 副标题}
> **评级：… | 目标价：… | 现价：… | 空间：… | 时间窗口：12 个月**

[门面三框·封面级,不另起 H2]
  · 评级框（评级/现价注日期/目标价/上行空间）
  · 基本数据表（总股本/总市值/52周区间/日均成交额/流通情况/主要上市地）
  · 盈利预测与估值摘要表（5 年:分部收入逐行+总营收/YoY/归母净利/YoY/EPS/P/E/P/B;多分部强制按分部拆,不许只一行总营收）

## 投资摘要
### 公司近况（Why Now）   ← 叙事入口,近期异动归因,可分阶段(股价%+驱动力)
### 核心判断              ← Key Takeaways:3-5 条 ■,每条加粗判断+关键数字,名词裸用不塞括号
```

门面三框是封面级浓缩框(回答"是什么/值多少/凭什么"),不是"把表格当独立 H2 章节",也不是"12 行技术面板大表灌进叙述"。数字接 VM(见 [估值章接法.md](./估值章接法.md))。门面之后正文直接从"一、公司概况"起(ICA 不写独立引言段)。

---

## 三、Frontmatter(11 项,必填)

```yaml
---
company:          # 公司全名 + 英文名
ticker:           # 代码.市场(002885.SZ / TSM (NYSE ADR))
date:             # YYYY-MM-DD
rating:           # BUY（买入）/ HOLD（持有）/ SELL（卖出）
target_price:     # 币种+值
current_price:    # 币种+值
upside:           # ±XX%
market_cap:       # 币种+值
sector:           # 行业/主题
analyst:          # 分析师姓名
tags:             # [equity-research, initiating-coverage, ...]
---
```

最常漏的 3 项:`market_cap` / `sector` / `analyst`,不能省。

---

## 四、体量(不分页,字数 + 图数)

ICA 的"25-45 页"是 DOCX 概念,Obsidian md 不分页。体量按字数 + 图数 + 表数,沿用 ICA 重量级:

| 维度 | 基准 |
|---|---|
| 正文字数 | 1.5-3 万字(火力集中在第三章壁垒 + 第四章估值) |
| ECharts 图 | 40-70 张(全 inline,图表 N 连续编号,见 [obsidian-echarts适配.md](./obsidian-echarts适配.md)) |
| 关键财务表 | 8-15 张 |
| 版面密度 | 60-80% 版面有内容;**图文交替穿插**(段→图→段→图),图由前文论证引出、紧跟其后,不堆"一坨文字 + 一串图"(见 echarts适配 §三) |

---

## 五、黑名单(Obsidian 绝对不出现)

- ❌ Obsidian callout(`> [!warning]` / `> [!note]` / `> [!example]` / `> [!info]` / `> [!tip]`)
- ❌ Emoji 装饰章节标题(`## 📋 投资摘要` → `## 投资摘要`;emoji 仅允许在图注 📊/📂)
- ❌ 英文章节标题(`## Investment Thesis` / `## Rating Box` → 全中文)
- ❌ Mermaid 流程图(用 ECharts sankey/scatter/tree 替代)
- ❌ DOCX 专属指令(page break / page number / 字体 / 字号 / 边框 / `w:rFonts` / `tblBorders`)
- ❌ 流程性元话语("本节由 agent 自动生成" / "数据生成于 YYYY" / "脚本路径" / "源数据见…"——过程脚手架不进交付物)
- ❌ 把表格当独立 H2 章节(门面三框除外)

## 六、白名单(保留)

- ✅ `■` 项目符号(核心判断 Key Takeaways 的 3-5 条 bullets 起头)
- ✅ 加粗 `**...**`(关键数字、评级、术语首次出现)
- ✅ `[[wiki link]]` 跨研报双链(每篇 ≥3 处)。**表格单元格内带显示名的双链必须转义管道符**:`[[页面\|显示名]]` 而非 `[[页面|显示名]]`(未转义的 `|` 会被当列分隔符把单元格劈成两列)
- ✅ 财务/行业术语保留英文:DCF/WACC/FCF/EBITDA/EPS/CAGR/Capex/TAM/SAM/SOM/Comps/Bull/Base/Bear/CoWoS/HBM/EUV/GAA/ASIC/GPU/TPU/SoC 等
- ✅ 评级双语:BUY(买入)/HOLD(持有)/SELL(卖出)
- ✅ `> 📂 资料来源` 一行来源注(图下只留来源;对图的解读写在**图前的正文段落**,不在图底写「📊 看图要点」。见 [obsidian-echarts适配.md](./obsidian-echarts适配.md) §三图文混排)

---

## 七、表格

- 纯 Markdown `|` 语法,不用 HTML/Word 表格。
- **多分部财务表**:年份为列(含 A/E 标识),指标为行;**收入按分部逐行拆 + 主分部占比 + 总营收行**。
- **份额/对手表**:`厂商 | 年份份额 | 排名 | 备注`;对手表 `竞争对手 | 定位与规模 | 核心优势 | 相对本公司弱势 | 市场份额`。
- **vs 卖方一致预期对标表**:`指标 | 本研究 | 卖方一致预期 | 差异及分歧根源`。
- 每张表表下注来源行(来源 + 日期)。
- 表格单元格内 `[[wiki link]]` 带显示名必须转义 `[[页面\|显示名]]`。

---

## 写完自检

- [ ] frontmatter 11 项齐全(`market_cap`/`sector`/`analyst` 不漏)
- [ ] 门面首页:主标题+副标题+评级条+三框(评级/基本数据/盈利预测摘要,多分部收入逐行)+ 投资摘要(Why Now + 核心判断 Key Takeaways)
- [ ] 全中文标题、判断式;无英文章节标题;无 callout;无 emoji 标题;无 Mermaid;无 DOCX 指令;无流程性元话语
- [ ] 表格不作为独立 H2(门面三框除外);表格内 wiki link 转义管道符
- [ ] 体量:1.5-3 万字、40-70 ECharts、8-15 财务表
- [ ] 跨研报双链 ≥3 处
- [ ] 年份 A/E 标注
