# valuation-model

将 `initiating-coverage` 的定性研报转为公式联动的估值模型。模型从业务驱动量逐级推导隐含股价，用历史财务和股价校准，并通过 P/E、P/B、EV/Sales、DCF 或 SOTP 中至少两种方法交叉核对。

取代旧 `coverage-model`(已退役留档)。核心改进:

1. **所有中间结果必须被下游公式引用**。旧版中 AIDC、营收和估值三张表互不引用,价格直接使用卖方盈利,本版需要消除这个问题。
2. **业务驱动量菜单**(扩展方式):结构公司无关,换一类公司只换最底层那一节基准(`references/anchors/`)。
3. **历史数据双职责**:历史财务依据投影基年;历史估值带校准目标倍数 + 现价 reality check。
4. **按步骤读取说明**:`SKILL.md` 说明执行顺序，具体方法放在对应 reference。

## 结构
```
SKILL.md                  执行步骤和验收要求
references/               每步细节(用到哪步读哪个)
  01-anchor-selection     选业务驱动量 + 估值方法(唯一卡点)
  02-data-grounding       4 源 + 历史财务 + 历史估值带
  03-modeling             连接公式 + 连通性自检 ★
  04-valuation            多种估值方法 + 逐年隐含价 + 历史带 overlay
  05-excel-format         视觉规范(从 coverage-model 搬)
  06-sourcing             溯源
  valuation-lenses        估值方法菜单
  anchors/                业务驱动量菜单(外延)
assets/quality-checklist  验收
scripts/                  recalc.py(跨平台重算)+validate_valuation.py(业务校验)+extract_model_tables.py
valuation_kit/            公共 Python 包和 JSON Schema
examples/                 通用入口和旧导入兼容层
```

## 状态
公开来源或用户材料先写入标准 `input.json`。`valuation_kit` 负责校验并生成 Excel，`scripts/recalc.py` 重算公式，`scripts/validate_valuation.py` 检查公式联动、来源、估值方法和历史回测。公司专属数据和脚本不放在仓库内。
