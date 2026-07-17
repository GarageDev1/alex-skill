---
name: valuation-model
description: 把 equity-research-obsidian 的定性研报转换为估值模型:从可观测的业务驱动量推导营收、利润和隐含股价;修改假设后,相关结果随公式更新。模型结合历史财务与股价校准,并用 PE/PB/DCF 中至少两种方法交叉核对,输出 Excel、决策 memo,再回填同 ticker 首次覆盖报告的估值章节。报告、journal 和决策统一引用本模型的隐含价。适用:已有定性研报,需要建立业务预测和估值模型;也被 equity-research-obsidian 的 Phase 3 按名调用。触发:给 ticker + "估值"/"建模"/"更新模型"/"接着研报做模型"。不适用:纯定性研究(用 equity-research-obsidian)、纯 DCF 模板(用 financial-analysis:dcf-model)。
---

> **语言**: 中文输出,专业术语保留英文 (DCF, TAM, P/E, P/B, ROE, EBITDA, HBM, CoWoS 等)。
> **读者**: Agent 读完本文与对应 reference 后,应能从研报搭建一个底层假设、财务预测与隐含股价公式联动的估值 Excel。

# Valuation Model

这个 skill 用公式连接业务驱动量、营收、利润、估值和隐含股价。每个中间结果都要被下游公式引用。

它接在 `initiating-coverage`(定性研报)之后,把研报中的业务判断转换为可计算、可用历史数据核对的估值模型。

---

## 模型用于预测,不用于拟合现价

模型用于预测合理价格并支持投资决策,不得为贴合现价或卖方共识而回调假设。

- **现价只用于对照**:现价和 TTM 倍数是当前市场数据,不是预测目标。目标倍数 × 前瞻每股得到的隐含价可以高于或低于现价,差额用于判断上行或下行空间。不得为了让隐含价接近现价而回调业务假设。
- **共识只作参照**:模型可以给出与市场主流不同的估值,但必须有事实和完整推导过程支撑。`卖方研报共识` sheet 是参照,不是目标值。
- **判断幅度**:
  - 假设过于保守,可能低估公司的增长空间。
  - 假设过于激进,可能高估公司的合理价格。
  - 在事实和推导能够支撑的范围内确定假设。最新数据若支持高于或低于共识、现价的结论,按数据更新模型。
- **持续跟踪并修正预测**:在 `综合判断仪表盘` sheet 记录可观测指标(基本面拐点、估值错位、催化剂和情绪确认,见 `05` 4c)。指标变化时,更新模型并重新预测。

> 这与"当前估值使用 TTM"(`04`)并不冲突。TTM 用于计算当前市场估值,避免用预测数据反向证明当前 P/E 合理;前瞻预测则根据业务数据和假设独立计算,结果可以高于或低于现价及共识。

---

## 所有中间结果必须被下游公式引用

```
[① 业务驱动量] → [② TAM/需求] → [③ 营收] → [④ 利润/账面/FCF] → [⑤ 估值方法×≥2] → [隐含价]
   ↑ 只有这一节随公司类型变化         ②~⑤ 的计算结构不变,且每一节的输入必须引用上一节的输出
```

- **不保留未被引用的 sheet**:计算结果没有下游公式引用时,删除该内容或补全公式关系。
- **不得用外部净利润替代模型结果**:营收模型完成后,利润和估值必须继续引用模型公式,不能直接填入卖方净利润。
- **验收测试**:修改最底层的假设(如 AIDC 装机 GW 或出货量)后,隐含股价必须随公式变化;否则公式链不合格。

> 旧 coverage-model 中,AIDC 拆解、营收预测和 PB 估值互不引用,隐含价格直接使用卖方盈利,导致业务模型不影响估值结果。详见 `references/03-modeling.md`。

---

## 建模精度分配 ★

这个 skill 与通用 DCF 的主要区别是:营收由装机 GW、出货量、供需平衡等业务驱动量自下而上推导,不直接填入收入增速。因此:

- **业务预测需要详细依据**:业务驱动量、TAM 和营收的推导需要逐项列出来源和计算公式。
- **利润和账面价值采用简化假设**:营收 × margin 假设得到盈利或 FCF。**不做三表勾稽、不做营运资本(DSO/DIO/DPO)、不逐科目拆 opex**。
- **估值倍数(PE/PB)→ 假设层**:基于倍数假设给隐含价,用三情景架构(情景切换 + 估值对比)看区间。
- **产出 = 方向性判断**(便宜 / 合理 / 已透支 N 年),服务大方向选择,**不追求精确到分的目标价**。

> 业务驱动量变化后,隐含价必须随公式更新。步骤 ④ 使用一项 margin 假设,不展开 20 行会计科目。

## 业务驱动量与历史数据

模型使用两类基础数据:

- **业务驱动量(最上游输入)**:使用可观测、可证伪且与公司需求直接相关的指标。指标随公司类型变化,见 `references/anchors/`。数据分为两部分:【行业数据】(产业链总规模和成本拆解,可供同一产业链公司复用,见 `shared-base/`)和【公司份额】(各公司的收入份额)。**同一产业链的公司必须引用同一份行业数据**,以保持口径一致。
- **历史数据**用于两项工作:
  - **历史财务(5yr 三表+分部)** → 确定步骤 ③④ 的预测基年。预测必须从真实历史数据开始,不得使用无法解释的 plug。
- **历史股价 + 历史估值带(P/E·P/B·EV/EBITDA·ROE 分位)** → ①校准步骤 ⑤ 的目标倍数,确认历史范围和当前分位;②将逐年隐含价与历史股价区间对照,判断现价超过历史上沿的幅度,以及估值区间是否已经发生变化。

> 连通性要求保证"链是通的";历史数据保证"链能复现现实、现价能被解释"。两个是不同验收维度,缺一不可。

---

## 6 步工作流(渐进式披露:用到哪步读哪个 reference,不要一次全读)

| 步 | 名称 | Reference | 暂停? |
|---|---|---|---|
| 1 | 接研报 + 选择业务驱动量 | `references/01-anchor-selection.md` | **默认不暂停**:决策树命中唯一解 → 自动选择业务驱动量+估值方法,理由写进备忘录;仅当歧义(菜单没有 / ≥2 候选都说得通)才停下让用户确定 |
| 2 | 数据落地 | `references/02-data-grounding.md` | 不暂停(公开来源/用户提供材料 + 历史数据 → 标准 `input.json`;外部数据 skill 仅作可选加速) |
| 3 | 建立计算关系 | `references/03-modeling.md` | 不暂停(连接公式并检查联动) |
| 4 | 使用多种方法估值 | `references/04-valuation.md` | 不暂停(PE/PB/DCF ≥2 估值方法 + 逐年隐含价 + 历史带 overlay) |
| 5 | Excel 输出 + 溯源 + 模型校验 | `references/05-excel-format.md` + `references/06-sourcing.md` | 不暂停(生成 → `scripts/recalc.py` 重算回写缓存值 → 溯源 → `scripts/validate_valuation.py` 校验;**verdict=FAIL 禁止交付**) |
| 6 | Reconcile + 决策 memo | `references/07-reconcile-memo.md` | 不暂停(对账同 ticker 首次覆盖报告 → 分歧表 + 决策 memo + 回填报告估值章节) |

**节奏**:默认 6 步连续完成、全程不暂停;只有第 1 步选择业务驱动量出现歧义、或 Step 5 校验 critical fail 时才停(异常路径)。基准 / 估值方法的选择理由写进「估值逻辑」sheet 与决策 memo,用户事后否决,否决即重新执行。

---

## 业务驱动量类型 → `references/anchors/`

| 公司类型 | 业务驱动量 | 例子 |
|---|---|---|
| 算力/半导体链 | AIDC 装机 GW × 单位 $ 含量 | Hynix / NVDA / TSMC |
| 消费电子 | 出货量 × ASP × 渗透率/换机周期 | Apple / 小米 |
| 平台 | MAU/DAU × ARPU × 时长 | Meta / 抖音 |
| 周期品 | 供需平衡表 → 价格 | 油 / 铜 / 内存价格 |
| 消费品牌 | 终端动销 + 提价 + 渗透 | 白酒 / 饮料 |
| SaaS | 席位 × ACV / 用量计费 | Snowflake |
| 浮存金/储备型 | 平均 float × 资产端 yield | Circle(USDC) / 保险浮存 / 货基 |

新增公司类型时,在菜单中增加一项;步骤 ②~⑤ 的计算结构和联动要求保持不变。

## 估值方法菜单 → `references/valuation-lenses.md`

重资产周期 = P/B-ROE(+重估期补 P/E) / 稳态 = DCF + P/E / 高增长无盈利 = EV/Sales + path / 银行险 = DDM·P/B-ROE·EV。
**强制 ≥2 估值方法交叉核对,禁止单估值方法下结论。**

---

## 验收 → `assets/quality-checklist.md`

- ☐ **连通性**:改底层假设,隐含价跟着动
- ☐ **公式引用完整**:没有未被下游公式引用的 sheet/行
- ☐ **基年真实**:③④ 投影起点 = 历史真实值(不许 plug)
- ☐ **倍数在带内**:目标倍数落在历史估值带内,突破要写明理由
- ☐ **现实检验**:逐年隐含价叠在历史股价带上,现价位置被解释
- ☐ **多种估值方法**:≥2 估值方法
- ☐ **溯源**:每个关键硬编码行有可见溯源(逻辑/备注列;**禁单元格批注**,见 `06`)
- ☐ **$B 口径统一(scale 要求)**:营收/利润/EBITDA/权益等财务量纲行**一律 USD $B**,不得用本币百万/亿/兆/EUR/SEK/NT$ 作主口径(市值可双币种、可附 FX 行=唯一例外)。`validate_valuation.py` 的「$B 口径统一」必须 ≠ fail。见 `references/02-data-grounding.md §29`
- ☐ **结构 = 业务驱动量驱动模型**:13~14 张标准 sheet(封面/历史财务/股价/共识/历史倍数/倍数假设/情景切换/业务驱动量/分部测算/利润假设/情景估值/估值对比/仪表盘),不得改用旧 `financial-analysis` 的 IS/BS/CF/DCF/Comps 模板。`validate_valuation.py` 的「结构」必须 ≠ fail
- ☐ **recalc 回写缓存值**:用 `scripts/recalc.py --backend auto` 重算并 Save；脚本按可用性尝试 Windows Excel COM、macOS Excel AppleScript、LibreOffice headless，落盘文件必须含缓存计算值:否则报告/memo 取数读到空、历史回测检查跑不了
- ☐ **模型校验**:`python scripts/validate_valuation.py <model.xlsx>` 返回 **verdict ≠ FAIL** 后方可交付;FAIL(尤其 scale / 结构 / 连通性)时返回 Step 3–5 修正

---

## 公共包维护规则

通用逻辑放在 `valuation_kit/`，公司数据只放在符合 Schema 的仓库外 `input.json`。仓库不再保存公司专属生成脚本。

| 改动属于 | 落点 |
|---|---|
| 标的数据或判断变化 | 更新该标的的外部 `input.json`，保留来源、日期和币种口径 |
| 通用工作表或估值方法变化 | 修改 `valuation_kit/`、对应 reference 和测试 |
| Excel 渲染或公式兼容问题 | 修改 `valuation_kit/styles.py` 或相关模块，并更新校验器 |

- **同步 = 替换,不是追加**:立新规则时,grep 全 skill 把被取代的旧规则(references / SKILL / checklist / examples / scripts 的所有出现处)删掉或改写,禁止新旧并存:否则文档变地层堆积,下个会话读到旧地层。
- **能机检的规则必须进 `validate_valuation.py`**:只写在 prose 里的规则靠下一个会话的自觉,迟早漂移;进了 validator 才算真正沉淀。
- **验收标准**:新标的只需要真实 `input.json` 和 `build_template.py`。生成后运行 `recalc.py` 与 `validate_valuation.py`，不得依赖旧公司脚本。

---

## 输出物(显式路径优先,`out/` 只是工作区)

`out/` 仅存中间版本与调试产物,**不是默认交付地**。最终目录按以下优先级解析:命令/调用方显式参数 > `VALUATION_OUTPUT_DIR` > 当前 workspace 的 `outputs/`。报告输入同理:显式参数 > `VALUATION_REPORT_PATH` > 不提供(此时跳过报告回填)。最终交付三件套落到解析后的输出目录:

- `<ticker>_valuation_model.xlsx` — 主交付物(底层假设与隐含价格公式联动)
- `<ticker>_input.json` — 数据 + 完整 source 元数据(**全研究的假设 SOT**:报告中可量化判断必须与它一致)
- `<ticker>_决策memo.md` — **强制产出**(Step 6):thesis 一句话 + 三情景隐含价 vs 现价 + 与报告/共识的分歧表 + 证伪触发器 + 评级。这是模型与报告的统一结论层,对接 investment-journal。

---

## 目录地图

```
valuation-model/
├── SKILL.md                          ← 本文件(薄主控)
├── references/
│   ├── 01-anchor-selection.md        选业务驱动量 + 估值方法(唯一卡点)
│   ├── 02-data-grounding.md          4 源 + 历史财务 + 历史股价&估值带
│   ├── 03-modeling.md                公式连接 + 联动检查 ★
│   ├── 04-valuation.md               多种估值方法 + 逐年隐含价 + 历史带 overlay
│   ├── 05-excel-format.md            视觉规范(从 coverage-model 搬)
│   ├── 06-sourcing.md                溯源 → 可见注释列(禁单元格批注)
│   ├── 07-reconcile-memo.md          Step 6:对账报告 + 决策 memo + 回填估值章节
│   ├── valuation-lenses.md           估值方法菜单
│   ├── lenses/dcf-lens.md            DCF 估值方法深化工艺(FCF 接主要计算关系/TV 双法/WACC/双敏感性)
│   └── anchors/                      各类业务驱动量与公司份额计算方法
├── shared-base/                      业务驱动量·共享基座层(产业链级总规模, 建一次反复用)
│   ├── README.md                     基座机制 + 维护规则
│   └── compute-aidc-base.{json,md}   AI 算力链基座(AIDC CapEx + $/GW 成本树)
├── assets/quality-checklist.md       验收清单
├── scripts/                          自包含工具(重算/校验/模型表抽取)
│   ├── validate_valuation.py         模型校验(连通性/孤儿/溯源/估值方法/**$B scale/结构**/历史回测)
│   ├── recalc.py                     跨平台重算回写缓存值(Excel COM / Excel AppleScript / LibreOffice)
│   └── extract_model_tables.py       从已 recalc 的模型抽『三情景估值汇总』等 Markdown 表(供报告嵌入)
├── valuation_kit/                    公共包：输入、工作表、情景、估值、币种和样式
└── examples/                         通用入口与旧导入兼容层
```
