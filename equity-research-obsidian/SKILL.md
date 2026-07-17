---
name: equity-research-obsidian
description: 生成首次覆盖报告及配套估值模型。报告写入 Obsidian,正文、篇幅和质量检查沿用 initial-coverage-advanced(ICA)规范;估值模型由 valuation-model 根据业务驱动量、营收、利润和估值公式计算隐含价。修改业务假设后,相关结果随公式更新。一次调用输出机构级首次覆盖研报、Excel 估值模型和决策 memo。报告估值章节统一读取模型结果。当用户说"做 initiating coverage""二级研究""深度拆解 XX 公司""initiate XX""写一份机构级研报""研报+模型一起做"时触发此 skill。
---

# Equity Research(ICA 的 Obsidian + 业务驱动量估值变体)

这个 skill 基于 [`initial-coverage-advanced`](../initial-coverage-advanced)(ICA)。报告正文结构、写作规范、体量和质检沿用 ICA;估值改由 `valuation-model` 完成,输出改为 Obsidian Markdown 和内联 ECharts。

一次调用,对一个 ticker 端到端产出**报告 + 模型三件套**,全部落盘到同一个已解析输出目录。输出目录优先级固定为:**显式 `output_dir` > 环境变量 `OBSIDIAN_VAULT_ROOT` > workspace fallback**:

```
<resolved_output_dir>/
├── {Company}.md                  ← 首次覆盖报告(中式四章 + 全 inline ECharts,1.5-3 万字 40-70 图)
├── {ticker}_valuation_model.xlsx ← 业务驱动量估值模型(valuation-model 产出)
├── {ticker}_input.json           ← 假设 SOT(数据 + source 元数据)
├── {ticker}_决策memo.md          ← 统一结论层(thesis + 隐含价 vs 现价 + 证伪触发器)
└── (中间产物:研究稿素材底稿,按需落盘)
```

**跨平台目录解析**:

1. 调用方显式给出 `output_dir` 时优先使用;绝对路径直接使用,相对路径相对于当前 workspace 根目录解析。
2. 未给 `output_dir` 且存在 `OBSIDIAN_VAULT_ROOT` 时,使用 `<OBSIDIAN_VAULT_ROOT>/research/公司研究/{Company}/`。
3. 两者都没有时,使用 `<workspace_root>/outputs/equity-research/{Company}/`;没有 Obsidian 也能运行。

先展开环境变量与 `~`,再用平台原生路径库(`pathlib.Path`)规范化;不得硬编码用户目录、盘符或路径分隔符。四件产物和中间产物始终复用同一个 `resolved_output_dir`。

**Skill 路径解析(不依赖 CWD)**:以本文件所在目录为 `ERO_SKILL_ROOT`;其父目录为 `REPO_ROOT`;进而得到 `ICA_SKILL_ROOT=REPO_ROOT/initial-coverage-advanced`、`VM_SKILL_ROOT=REPO_ROOT/valuation-model`。所有兄弟 Skill 文件和脚本均从这些根目录解析,不得假设命令从 `REPO_ROOT` 启动。

**默认语言**:简体中文,术语保留英文(EBITDA、EPS、DCF、WACC、CoWoS、HBM 等)。

---

## 文件地图(按执行工序加载,别开工全 Read)

本 skill 沿用 ICA 的写作、结构和质检要求,另行定义估值接入、图表和格式。按执行步骤读取标有 `[步N]` 的文件,不要在开始时一次读完所有文件。全程使用下方《核心原则》和《执行工序》。

**沿用 ICA**

- `[步1]` **ICA [`assets/研报写作规范.md`](../initial-coverage-advanced/assets/研报写作规范.md)** — 写作真相源(去 AI 味、判断式标题、数字优先、A/E 口径、枚举拆 bullet/推演用散文、逻辑链追归因)
- `[步1]` **ICA [`references/阶段1-公司研究.md`](../initial-coverage-advanced/references/阶段1-公司研究.md)** — 研究工序(投资逻辑链组织 / 分部拆分 / 一手披露)
- `[步4]` **ICA [`assets/报告模板与版式.md`](../initial-coverage-advanced/assets/报告模板与版式.md)** — 正文结构真相源(中式四章结构:一公司概况 / 二行业空间 / 三核心驱动·壁垒[最重] / 四盈利预测与估值+风险;封面;体量基准 1.5-3 万字 40-70 图 8-15 财务表)
- `[步4]` **ICA [`references/阶段5-报告组装.md`](../initial-coverage-advanced/references/阶段5-报告组装.md)** — 组装方法(结构取材 / 跨文件一致性 / vs 卖方共识+GAP% 强制嵌入)。DOCX 格式部分换成 Obsidian md
- `[步4]` **ICA [`assets/质量检查清单.md`](../initial-coverage-advanced/assets/质量检查清单.md)** — 交付前质检(DOCX 专属项由本 skill 格式覆盖层屏蔽)

**本 skill 自写(相对 ICA 的差异处)**

- `[步2]` **[`references/估值章节数据接入.md`](./references/估值章节数据接入.md)** — 接 valuation-model:recalc+validate 模型校验 / extract_model_tables 四表强制嵌入 / 业务驱动分析 / P/E 历史 vs 预测口径
- `[步3]` **[`references/obsidian-echarts适配.md`](./references/obsidian-echarts适配.md)** — 全 inline ECharts:40-70 图工程 / 语义配色 / 图型对照 / 模型语义图 / 图文混排
- `[步4]` **[`references/obsidian-格式覆盖.md`](./references/obsidian-格式覆盖.md)** — Obsidian 格式:frontmatter 11 项 / wiki link / 不分页 / callout 禁 / 门面首页

---

## 核心原则(三条)

1. **估值结论取自模型**:目标价、评级和三种情景统一使用 `valuation-model` 的逐年隐含价。同一 ticker 只保留一套估值数字,报告不得另算加权目标价。先完成 Excel 模型并通过校验,再从模型读取报告结论;不得先定结论再调整模型假设。
2. **模型的信息量必须进报告**:业务驱动关系、三情景假设、隐含价 vs 历史估值带、可跟踪仪表盘指标:这些是第四章正文素材,不能只留在 Excel 里不进报告。
3. **报告 ↔ 模型同源,单向耦合**:报告中的可量化判断必须与 `input.json` 一致;如有不一致,回到模型修正,不在报告中另行填数。模型是估值数字的唯一来源,不从报告回填数据。

---

## 执行要求摘要

1. **正文 = 中式四章**(ICA 报告模板):一公司概况 / 二行业空间 / 三核心驱动·壁垒(最重)/ 四盈利预测与估值+风险。判断式标题,顺序可按公司调整。
2. **判断式标题**:每个 H1/H2/H3 带判断,非中性名词("CAPEX 高度集中于四家超大厂" 优于 "CAPEX 总量")。
3. **去 AI 味**:删填充套话(值得注意的是/此外/然而)、砍 -ing 伪分析尾巴(彰显/体现/奠定基础)、不堆三段排比、不滥用破折号、无翻译腔黑话(成本桥/顺风逆风/敞口/硬约束)。详见 ICA 研报写作规范.md。
4. **A/E 年份口径**:历史年标 A、预测年标 E,盈利预测表分清。
5. **Key Takeaways**:报告开头摘要 3-5 条加粗判断领起,每条带关键数字(不是 topic 词)。
6. **数字同源**:估值/财务数字只能由 `<VM_SKILL_ROOT>/scripts/extract_model_tables.py` 从 VM 抽取后嵌入,**绝不手敲**。
7. **全 inline ECharts + 图文混排**:40-70 张,每张不写 `title.text`、用研报语义配色。图下**只一行资料来源**(无「看图要点」);对图的解读写进**图前的正文段落**,图由前文论证引出、紧跟其后,段→图交替穿插,不堆"一坨文字 + 一串图"。不用 PNG。

---

## 执行工序

一个 ticker 触发,连续推进:

0. **解析根目录** → 按上文规则解析 `resolved_output_dir`、`REPO_ROOT`、`ICA_SKILL_ROOT`、`VM_SKILL_ROOT`,校验三个同仓 Skill 目录与 VM 三个脚本存在
1. **研究素材** → 先读 ICA 阶段1 + 研报写作规范;按投资逻辑链组织、分部拆分、一手披露强制,产物 .md 研究稿
2. **估值建模** → 先读 [估值章节数据接入.md](./references/估值章节数据接入.md);按名调用同仓 `valuation-model` skill 执行 6 个步骤;**模型校验**:`<VM_SKILL_ROOT>/scripts/recalc.py` + `<VM_SKILL_ROOT>/scripts/validate_valuation.py`(verdict=PASS)
3. **图表** → 先读 [obsidian-echarts适配.md](./references/obsidian-echarts适配.md);每张图内联完整 layout 的 ECharts(40-70 张),不依赖外部自动补全机制
4. **组装报告** → 先读 ICA 阶段5 + [格式覆盖.md](./references/obsidian-格式覆盖.md) + ICA 质检清单;按结构/取材/跨文件一致性方法,产物 Obsidian md

每步产物先落盘,下一步读磁盘校验前置;管线可断点续跑。

**估值定稿前强制确认一次**:模型完成后,在第四章确定估值、目标价和评级前,请用户审核估值假设与结论。`valuation-model` 选择业务驱动量时如有多个合理选项,也可暂停确认。其余步骤连续推进。

只交付三件套 + 决策 memo,不造"完成总结/下一步说明"等多余文档。

---

## 质量与成功标准

完成标准:研报、估值模型和决策 memo 均已生成;报告估值数字与模型一致;第四章说明业务驱动量如何影响收入和估值;决策 memo 可直接提供给 `investment-journal` 建立 judgment。报告篇幅为 1.5-3 万字,包含 40-70 张 ECharts 图和 8-15 张财务表。交付前执行 ICA [`assets/质量检查清单.md`](../initial-coverage-advanced/assets/质量检查清单.md)(DOCX 项按格式覆盖层排除)和本 skill [估值章节数据接入.md](./references/估值章节数据接入.md) 的数字一致性检查。
