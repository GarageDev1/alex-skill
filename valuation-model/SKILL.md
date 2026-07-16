---
name: valuation-model
description: 把 equity-research-obsidian 的定性研报接成估值模型:从业务最底层的物理驱动量(物理锚)用公式一路推到隐含股价,改底层假设、隐含价跟着变;叠加历史财务与股价做校准,PE/PB/DCF 至少两种方法交叉验证,输出机构级 Excel + 决策 memo,并回填同 ticker 首次覆盖报告的估值章节。估值数字一律以本模型为准:报告、journal、决策都引用它的隐含价。适用:已有定性研报、要把业务和价格联动起来建估值模型;也被 equity-research-obsidian 的 Phase 3 按名内嵌调用。触发:给 ticker + "估值"/"建模"/"更新模型"/"接着研报做模型"。不适用:纯定性研究(用 equity-research-obsidian)、纯 DCF 模板(用 financial-analysis:dcf-model)。
---

> **语言**: 中文输出,专业术语保留英文 (DCF, TAM, P/E, P/B, ROE, EBITDA, HBM, CoWoS 等)。
> **读者**: Agent 读完本文与对应 reference 后,应能从研报搭建一个底层假设、财务预测与隐含股价公式联动的估值 Excel。

# Valuation Model

这个 skill 用公式连接业务驱动量、营收、利润、估值和隐含股价。每个中间结果都要被下游公式引用。

它接在 `initiating-coverage`(定性研报)之后:研报告诉你公司是干啥的,这个 skill 把它折成一个能算出隐含股价、并且能对回历史现实的模型。

---

## 模型用于预测,不用于拟合现价

模型用于预测合理价格并支持投资决策,不得为贴合现价或卖方共识而回调假设。

- **不是拟合现价**:当下锚(现价 / TTM 倍数)只是 reality-check 的**观测点**(知道"我们现在在哪"),不是**预测的目标**。预测(目标倍数 × 前瞻每股 → 隐含价)**应当、而且往往会**和现价不同——那个差,就是投资信号(上行 / 下行空间)。**为了让隐含价"贴近现价"去回调假设,等于放弃模型的预测功能。**
- **共识只作参照**:模型可以给出与市场主流不同的估值,但必须有事实和完整推导过程支撑。`卖方研报共识` sheet 是参照,不是目标值。
- **判断幅度**:
  - 拍**太保守** → 错过一只很有潜力的票(本该买、没买)。
  - 拍**太激进** → 买了就亏。
  - 在事实和推导能够支撑的范围内确定假设。最新数据若支持高于或低于共识、现价的结论,按数据更新模型。
- **持续跟踪并修正预测**:在 `综合判断仪表盘` sheet 记录可观测指标(基本面拐点、估值错位、催化剂和情绪确认,见 `05` 4c)。指标变化时,更新模型并重新预测。

> 和"当下锚拍 TTM"(`04`)**不矛盾**:TTM 锚是"现在贵不贵"的**事实观测**(防止用自己的乐观预测去循环论证当前 P/E);**预测**是另一回事——它就该独立地、基于事实+逻辑,给出一个可能高于、也可能低于现价 / 共识的方向性判断。两者是"我们在哪" vs "该往哪走",分开。

---

## 所有中间结果必须被下游公式引用

```
[① 物理锚] → [② TAM/需求] → [③ 营收] → [④ 利润/账面/FCF] → [⑤ 估值镜头×≥2] → [隐含价]
   ↑ 只有这一节随公司类型换          ②~⑤ 骨架不变, 且每一节的输入必须来自上一节的输出(公式引用)
```

- **不保留未被引用的 sheet**:计算结果没有下游公式引用时,删除该内容或补全公式关系。
- **不许并行抄外部盈利截断链路**:不能营收模型辛苦建完,估值却直接硬塞一个卖方净利润,把链截断。
- **验收测试**:修改最底层的假设(如 AIDC 装机 GW 或出货量)后,隐含股价必须随公式变化;否则公式链不合格。

> 旧 coverage-model 中,AIDC 拆解、营收预测和 PB 估值互不引用,隐含价格直接使用卖方盈利,导致业务模型不影响估值结果。详见 `references/03-modeling.md`。

---

## 模型海拔:严谨度花在物理驱动,不花在会计 ★

这个 skill 相比通用 DCF 的**唯一卖点**:营收是**从底层业务的物理量 bottom-up 驱动**出来的(装机 GW / 出货量 / 供需平衡),不是拍一个增速。所以:

- **严谨度预算 → 链条上半段**(物理锚 → TAM → 营收)。这里多花功夫、多要出处。
- **利润/账面 → 粗颗粒假设**:营收 × 一个 margin 假设 → 大致盈利/FCF。**不做三表勾稽、不做营运资本(DSO/DIO/DPO)、不逐科目拆 opex**。会计精度不是目标。
- **估值倍数(PE/PB)→ 假设层**:基于倍数假设给隐含价,用三情景架构(情景切换 + 估值对比)看区间。
- **产出 = 方向性判断**(便宜 / 合理 / 已透支 N 年),服务大方向选择,**不追求精确到分的目标价**。

> 连通性铁律照旧:链要通、物理锚一动隐含价要动。只是中间 ④ 是**一根 margin 假设线**,不是 20 行会计。

## 双根基:物理锚 + 历史底座

链不是凭空起的,两个根基撑着它:

- **物理锚(链的最底层输入)**:公司需求最底层、可观测、可证伪的物理量。随公司类型换,见 `references/anchors/`。物理锚本身是**两层**——【共享基座】(产业链级总盘子+成本拆解,建一次反复用,见 `shared-base/`)+【公司切片】(每家吃哪一刀,现做)。**同一产业链的公司必须引用同一个基座**,模型才互相一致。
- **历史底座(链的现实校准器)**,干两件别的环节替代不了的事:
  - **历史财务(5yr 三表+分部)** → 锚定 ③④ 投影的**基年**(投影必须从真实历史接出来,不许 plug/漂浮)。
  - **历史股价 + 历史估值带(P/E·P/B·EV/EBITDA·ROE 分位)** → ①校准 ⑤ 的目标倍数(倍数历史上真出现过吗、现在在什么分位)②把隐含价阶梯叠回历史股价带做 reality check(现价是历史上沿之上,还是进了新 regime)。

> 连通性铁律保证"链是通的";历史底座保证"链能复现现实、现价能被解释"。两个是不同验收维度,缺一不可。

---

## 6 步工作流(渐进式披露:用到哪步读哪个 reference,不要一次全读)

| 步 | 名称 | Reference | 暂停? |
|---|---|---|---|
| 1 | 接研报 + 选锚 | `references/01-anchor-selection.md` | **默认不暂停**:决策树命中唯一解 → 自动选锚+镜头,理由写进备忘录;仅当歧义(菜单没有 / ≥2 候选都说得通)才停下让用户拍板 |
| 2 | 数据落地 | `references/02-data-grounding.md` | 不暂停(公开来源/用户提供材料 + 历史底座 → 标准 `input.json`;外部数据 skill 仅作可选加速) |
| 3 | 建模通链 | `references/03-modeling.md` | 不暂停(接通骨架 + 连通性自检) |
| 4 | 多镜头估值 | `references/04-valuation.md` | 不暂停(PE/PB/DCF ≥2 镜头 + 隐含价阶梯 + 历史带 overlay) |
| 5 | Excel 输出 + 溯源 + 校验门 | `references/05-excel-format.md` + `references/06-sourcing.md` | 不暂停(生成 → `scripts/recalc.py` 重算回写缓存值 → 溯源 → `scripts/validate_valuation.py` 校验;**verdict=FAIL 禁止交付**) |
| 6 | Reconcile + 决策 memo | `references/07-reconcile-memo.md` | 不暂停(对账同 ticker 首次覆盖报告 → 分歧表 + 决策 memo + 回填报告估值章节) |

**节奏**:默认 6 步一气跑完、全程不暂停;只有第 1 步选锚出现歧义、或 Step 5 校验 critical fail 时才停(异常路径)。锚 / 镜头的选择理由写进「估值逻辑」sheet 与决策 memo,用户事后否决,否决即重跑。

---

## 物理锚菜单(外延机制) → `references/anchors/`

| 公司类型 | 物理锚 | 例子 |
|---|---|---|
| 算力/半导体链 | AIDC 装机 GW × 单位 $ 含量 | Hynix / NVDA / TSMC |
| 消费电子 | 出货量 × ASP × 渗透率/换机周期 | Apple / 小米 |
| 平台 | MAU/DAU × ARPU × 时长 | Meta / 抖音 |
| 周期品 | 供需平衡表 → 价格 | 油 / 铜 / 内存价格 |
| 消费品牌 | 终端动销 + 提价 + 渗透 | 白酒 / 饮料 |
| SaaS | 席位 × ACV / 用量计费 | Snowflake |
| 浮存金/储备型 | 平均 float × 资产端 yield | Circle(USDC) / 保险浮存 / 货基 |

**新公司类型 = 往菜单加一条,骨架(②~⑤)和连通性铁律一个字不动。**

## 估值镜头菜单 → `references/valuation-lenses.md`

重资产周期 = P/B-ROE(+重估期补 P/E) / 稳态 = DCF + P/E / 高增长无盈利 = EV/Sales + path / 银行险 = DDM·P/B-ROE·EV。
**强制 ≥2 镜头三角验证,禁止单镜头下结论。**

---

## 验收 → `assets/quality-checklist.md`

- ☐ **连通性**:改底层假设,隐含价跟着动
- ☐ **无孤岛**:没有算了却没人引用的 sheet/行
- ☐ **基年真实**:③④ 投影起点 = 历史真实值(不许 plug)
- ☐ **倍数在带内**:目标倍数落在历史估值带内,突破要写明理由
- ☐ **现实检验**:隐含价阶梯叠在历史股价带上,现价位置被解释
- ☐ **多镜头**:≥2 估值镜头
- ☐ **溯源**:每个关键硬编码行有可见溯源(逻辑/备注列;**禁单元格批注**,见 `06`)
- ☐ **$B 口径统一(scale 铁律)**:营收/利润/EBITDA/权益等财务量纲行**一律 USD $B**,不得用本币百万/亿/兆/EUR/SEK/NT$ 作主口径(市值可双币种、可附 FX 行=唯一例外)。`validate_valuation.py` 的「$B 口径统一」必须 ≠ fail。见 `references/02-data-grounding.md §29`
- ☐ **结构 = 物理锚通电模型**:13~14 张标准 sheet(封面/历史财务/股价/共识/历史倍数/倍数假设/情景切换/物理锚/分部测算/利润假设/情景估值/估值对比/仪表盘),**绝不退回旧 financial-analysis 的 IS/BS/CF/DCF/Comps 模板**。`validate_valuation.py` 的「结构」必须 ≠ fail
- ☐ **recalc 回写缓存值**:用 `scripts/recalc.py --backend auto` 重算并 Save；脚本按可用性尝试 Windows Excel COM、macOS Excel AppleScript、LibreOffice headless，落盘文件必须含缓存计算值——否则报告/memo 取数读到空、历史回测检查跑不了
- ☐ **校验门**:`python scripts/validate_valuation.py <model.xlsx>` 跑出 **verdict ≠ FAIL** 方可交付;FAIL(尤其 scale / 结构 / 连通性)回 Step 3–5 重修,不许带病落盘

---

## Kit-first 同步规则

**核心一句话:通用逻辑禁止首次落地在公司专用脚本里。**

对实例模型(`examples/build_<ticker>.py`)的任何修改,commit 前先分类、对号入座:

| 改动属于 | 落点 |
|---|---|
| ① 标的自己的数据/判断变了(新季报、新研报、改情景取值) | 只动实例脚本 |
| ② 通用方法论变了(新 sheet 架构 / 新铁律 / 新检查) | **先改 reference + `examples/build_kit.py` 骨架函数,实例只能调用 kit** |
| ③ Excel 工程坑(渲染 / locale / 公式兼容) | 修进 kit helper + `scripts/validate_valuation.py` |

- 发现自己正在给 `build_<ticker>.py` 写一段"下家公司也会用"的代码 = 停,先抽进 kit,再让实例调用。
- **同步 = 替换,不是追加**:立新规则时,grep 全 skill 把被取代的旧规则(references / SKILL / checklist / examples / scripts 的所有出现处)删掉或改写,禁止新旧并存——否则文档变地层堆积,下个会话读到旧地层。
- **能机检的规则必须进 `validate_valuation.py`**:只写在 prose 里的规则靠下一个会话的自觉,迟早漂移;进了 validator 才算真正沉淀。
- **验收标准**:新标的建模从 kit + `build_template.py` 出发,**全程不打开任何旧公司的 build 脚本**,产物过 validate。途中任何"不得不去翻旧公司脚本"的瞬间 = 欠的债,按本协议还进 kit/reference。

---

## 输出物(显式路径优先,`out/` 只是工作区)

`out/` 仅存中间版本与调试产物,**不是默认交付地**。最终目录按以下优先级解析:命令/调用方显式参数 > `VALUATION_OUTPUT_DIR` > 当前 workspace 的 `outputs/`。报告输入同理:显式参数 > `VALUATION_REPORT_PATH` > 不提供(此时跳过报告回填)。最终交付三件套落到解析后的输出目录:

- `<ticker>_valuation_model.xlsx` — 主交付物(端到端通电模型)
- `<ticker>_input.json` — 数据 + 完整 source 元数据(**全研究的假设 SOT**:报告中可量化判断必须与它一致)
- `<ticker>_决策memo.md` — **强制产出**(Step 6):thesis 一句话 + 三情景隐含价 vs 现价 + 与报告/共识的分歧表 + 证伪触发器 + 评级。这是模型与报告的统一结论层,对接 investment-journal。

---

## 目录地图

```
valuation-model/
├── SKILL.md                          ← 本文件(薄主控)
├── references/
│   ├── 01-anchor-selection.md        选物理锚 + 估值镜头(唯一卡点)
│   ├── 02-data-grounding.md          4 源 + 历史财务 + 历史股价&估值带
│   ├── 03-modeling.md                通链 + 连通性自检 ★
│   ├── 04-valuation.md               多镜头 + 隐含价阶梯 + 历史带 overlay
│   ├── 05-excel-format.md            视觉规范(从 coverage-model 搬)
│   ├── 06-sourcing.md                溯源 → 可见注释列(禁单元格批注)
│   ├── 07-reconcile-memo.md          Step 6:对账报告 + 决策 memo + 回填估值章节
│   ├── valuation-lenses.md           估值镜头菜单
│   ├── lenses/dcf-lens.md            DCF 镜头深化工艺(FCF 接主链/TV 双法/WACC/双敏感性)
│   └── anchors/                      物理锚菜单·公司切片层(外延机制)
├── shared-base/                      物理锚·共享基座层(产业链级总盘子, 建一次反复用)
│   ├── README.md                     基座机制 + 维护规则
│   └── compute-aidc-base.{json,md}   AI 算力链基座(AIDC CapEx + $/GW 成本树)
├── assets/quality-checklist.md       验收清单
├── scripts/                          自包含工具(重算/校验/模型表抽取)
│   ├── validate_valuation.py         业务校验门(连通性/孤儿/溯源/镜头/**$B scale/结构**/历史回测)
│   ├── recalc.py                     跨平台重算回写缓存值(Excel COM / Excel AppleScript / LibreOffice)
│   └── extract_model_tables.py       从已 recalc 的模型抽『三情景估值汇总』等 Markdown 表(供报告嵌入)
└── examples/                         Hynix worked example(样板)
```
