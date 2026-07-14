# 物理锚:算力/半导体链(AIDC × 单位 $ 含量)

> 适用:HBM/DRAM/NAND、GPU、代工、设备、网络、电源、冷却 —— 需求由 AI 数据中心资本支出驱动的公司。
> 范例:SK Hynix、NVDA、TSMC。

## ★ 这个锚是两层:基座(引用)+ 切片(现做)

```
物理锚 = 【共享基座】 + 【公司切片】

Layer 1 共享基座  → 别重推, 直接引用 ../../shared-base/compute-aidc-base.json
   ① 全球 AIDC CapEx 总额($/年, 2022A–2036E, 带源)
   ② $51.6B/GW 成本树 → 各部件占总 capex 的 %

Layer 2 公司切片  → 每家现做(本文件的活儿)
   ③ 这家吃哪几个 tree/chip 节点 (NVDA=GPU+网络 / Hynix=HBM+DRAM+NAND / TSMC=逻辑代工)
   ④ 切片 TAM = AIDC总额 × 部件占比%  → × 份额 → 营收
```

**第一步永远是先读基座**,不要再去网上重新搜一遍 AIDC 总量——那会和别的算力链公司模型对不上(Hynix 旧模型的真实翻车)。基座过时了(每季财报后)就更新 `shared-base/`,而不是在单个公司模型里另起炉灶。

## 链 build pattern(骨架)

```
① 基座:    AIDC CapEx 总额($/年)        ← 读 base.json: aidc_capex_forecast
② 切片占比: 公司部件占总 capex 的 %        ← 读 base.json: cost_breakdown_per_gw
③ 切片营收: AI直驱段 = AIDC CapEx × 收入强度; 周期段 = 上年 × 量 × 价
            (供需缺口用于价格路径/可行性旁注,不作单公司收入主驱动)
④ 利润:    营收 → 毛利率(节点/产品 mix)→ opex → 净利 → 账面
```

## 落进 Excel:基座必须接成 TAM 的来源(治旧版孤儿病)

旧 Hynix 模型把 AIDC 拆解算成孤岛、TAM 另抄 HSBC,两边不通电。**正确做法**:
- 把基座(①②)落成 Excel 里一张「产业链基座」sheet,数值来源写进可见备注/逻辑列。
- `②TAM` 的公式**引用**这张 sheet 的 capex 总额 × 部件占比 —— 这样改 AIDC capex,TAM→营收→股价全链动。

## ★ 收入拆法:把"能挂 AIDC capex 的"显式挂上,其余走周期(细拆 + AI 关联)
目的是**细拆 + 把 AI 驱动的部分都关联到 AIDC capex**,且每个数写 build-up 推导(见 `06`),不是拍。
内存公司收入拆成 **AI 直驱段(挂 capex)** + **消费段(周期)** 两类:

| 段 | 驱动方式 | build-up 逻辑(写进逻辑列) |
|---|---|---|
| **HBM** | `收入 = AIDC capex × HBM强度` | 强度 = 该公司当年 HBM 收入 ÷ AIDC capex,**锚最近实际年**(如 2025 ≈ 5.7%),随 HBM 渗透/含量升 → 后随竞争 normalize |
| **服务器常规 DRAM(DDR5)** | `收入 = AIDC capex × 服务器DRAM强度` | bit 随 AI 服务器出货 × 单台 DDR5 含量(GS:server DRAM +39%@26);服务器随 AIDC capex 放量 → 同样挂 capex |
| **企业级 eSSD(NAND)** | `收入 = AIDC capex × eSSD强度` | AI 机架存储(GS:Rubin 16TB/GPU+512GB/DPU × 72-GPU 机架 ≈ 1,161TB/机架 → +29EB@26);KV-cache/ICMSP 驱动 → 挂 capex |
| **消费 DRAM/NAND(Mobile/PC)** | 周期:`上年 ×(1+bit)×(1+价)` | bit 随手机/PC 出货(平/降,涨价致单机含量下调);**不挂 capex** |

**关键:强度全部锚最近实际年**(收入 ÷ capex 反推一个比值),再前后外推——这样只锚 1 个真实数,避免"产能×单片产出×ASP×份额"四个粗数连乘的失真(旧版软肋)。**不再用供给/需求取 min 作主驱动**(那是行业价格建模;单公司收入用 capex×强度更干净)。

### ★★ 强度反推的三条铁律(血泪)
1. **历史强度必须由实数公式反推,不许拍**:每个 AI 直驱段,历史列 = `历史段收入$B ÷ 当年 capex$B`(Excel 公式,不是输入一个数)。2025(最近实际年)算出来多少,就是锚。**前瞻必须从这个锚平滑续上,不准跳断层**——实测翻车:历史反推 4.1%,前瞻却拍 6.0%,凭空跳一截,不成立。前瞻第一年与锚的差,必须有 build-up 解释(如"该段约翻倍 > capex 1.7x 增速→强度小升")。
2. **单位校验(同币种÷同币种)**:强度 = 段收入$B ÷ capex$B,**两边必须同币种**。别把"28兆KRW"当成"$28B"去除 capex——这正是 4.1% 被误写成 5.7% 的 bug(28兆=$20B@1397,$20B÷$488B=4.1%)。反推前先把段收入折成 $B。
3. **纯 AI 段 vs 含基底段,历史可锚性不同**:
   - **纯 AI 驱动段(如 HBM)**:AI 爆发前几乎为零,收入÷capex 历史就干净(分子分母都小),2021 起逐年反推都成立。
   - **含非 AI 基底段(如服务器 DRAM 含云基底、eSSD 含消费 NAND 基底)**:早年(AI capex 极小)**大基底 ÷ 极小 AI capex → 比值灌爆失真**(服务器 DRAM 2021 反推出 50% 是假数)。处理:**早年(AI capex 未成规模的 2021-23)标 `n.m.`,只锚 AI capex 已成规模的近年(2024-25)**,前瞻从近年锚续。要在逻辑列写明"早年 n.m. 因 AI capex 前是云/消费基底,÷极小 capex 失真"。

## 价格:走 S/D 缺口逻辑(AI 通过缺口间接驱动全部内存价)
量挂 capex 后,**价格单独叠一层周期**:AI 吸走产能 → 全行业缺口(GS:DRAM -4.9%/NAND -4.2%@26)→ 价暴涨(2026 +112~176%/+137%)→ 2028-30 normalize(卖方不覆盖,explicit 标"本模型周期判断")。价格分歧大给区间(GS conventional +176% vs Citi blended +112%,口径差)。这条把"价"也间接挂上 AI。

## 产能 = fab 周期(仅作可行性检查,不作主驱动)
★ 注意:**产能/供给本质由建厂周期驱动**(洁净室/TSV 封装线/TC bonder 设备到位,1.5-2 年 lead time、新厂 ramp 时点),**不是 AIDC capex 的倍数**(旧版"产能=capex×强度"是为强行连通的 hack,概念错,已废)。若要核"公司产能够不够供它这份额",用**管理层 capex 指引/fab 产能规划**做一个旁注检查,不挂主链。

## 两条口径纪律(从基座继承,建模时别犯)

1. **护栏(spend tree)**:同一 tree 节点被多家声称,各家隐含 $/GW 之和 ≤ 节点 $/GW。
2. **resale 叠加**:HBM 的钱已包在 GPU 收入里——Hynix 与 NVDA 营收**不能相加**对总额,各对各自所在层。

## worked example
`../../examples/` 里的 SK Hynix 模型是这个锚的最新实作参考:切片拆成 HBM/服务器DRAM/eSSD 等 AI 直驱段 + 消费周期段,AI 段用 `AIDC CapEx × 收入强度`,周期段单独拍量价,并用 P/B + P/E 双镜头跑护栏。**例子只看结构,不抄数字。**

## TSMC:在 GPU 节点内切代工层
GPU 节点($26.25/GW)是含 HBM + 设计 + 代工的整颗售价;TSMC 只吃**代工**那层:
`TSMC SAM/GW = GPU/GW × (代工价 / GPU ASP)`,代工价占 GPU ASP 比例从 TSMC 对 NVDA 单价 / NVDA GPU ASP 反推。
