# valuation-model

把 `initiating-coverage` 的定性研报,接成一个**端到端通电**的估值模型——从业务最底层的物理驱动量(物理锚)推到隐含股价,叠历史财务/股价底座做现实校准,PE/PB/DCF 多镜头三角验证,输出机构级 Excel。

取代旧 `coverage-model`(已退役留档)。核心改进:

1. **链路必须端到端通电**(唯一铁律)——治旧版"AIDC/营收/估值三张表互不通电、价格架在卖方盈利上"的真实翻车。
2. **物理锚菜单**(外延机制)——骨架公司无关,换一类公司只换最底层那一节锚(`references/anchors/`)。
3. **历史底座双职责**——历史财务锚定投影基年;历史估值带校准目标倍数 + 现价 reality check。
4. **渐进式披露**——薄 SKILL.md(主控)+ references 按需加载,对标官方 `initiating-coverage`。

## 结构
```
SKILL.md                  薄主控(整体步骤 + 唯一铁律 + 验收)
references/               每步细节(用到哪步读哪个)
  01-anchor-selection     选物理锚 + 估值镜头(唯一卡点)
  02-data-grounding       4 源 + 历史财务 + 历史估值带
  03-modeling             通链 + 连通性自检 ★
  04-valuation            多镜头 + 隐含价阶梯 + 历史带 overlay
  05-excel-format         视觉规范(从 coverage-model 搬)
  06-sourcing             溯源
  valuation-lenses        估值镜头菜单
  anchors/                物理锚菜单(外延)
assets/quality-checklist  验收
scripts/                  recalc.py(跨平台重算)+validate_valuation.py(业务校验)+extract_model_tables.py
examples/                 Hynix worked example
```

## 状态
🟢 可用。6 步 references + 7 类 anchor build pattern + shared-base(算力链)+ build kit + 跨平台 recalc + validate_valuation.py(连通性/孤儿/溯源/镜头校验)均已落地。
本仓自包含执行链:公开来源或用户提供材料先落标准 `input.json`;Excel 由 `examples/build_kit.py` + `examples/build_template.py` 构建;公式缓存由 `scripts/recalc.py` 重算;最后由 `scripts/validate_valuation.py` 校验。外部数据/检索 skill 可以加速取数,但不是必需依赖。
待打磨:④ 营收→账面的逐科目 how-to 深度、各 anchor 真实 worked example。
