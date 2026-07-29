# 步骤 3：创建 JSON

先读取 `references/schema.md` 了解字段定义和约束。

## 观点卡

每章使用 2-5 张观点卡，每张卡只表达一个判断：**加粗标签 + 一句话 + 关键事实**。

## 表格

需要承载口径对照、业务拆分、竞品比较或情景假设时，使用 `blocks` 中的紧凑表格：
- 每表 2-4 列，超出时拆为两张表或两段。
- 不得用小字号直接填入。

## Dual 模式的 audience 标注

为每个章节、观点卡和视觉块显式写入 `audience`：
- `both`：允许出现在 Internal 和 External。
- `internal`：仅允许出现在 Internal。

External 采用"白名单优先"：未标为 `both` 的内容一律不进入外部版。

### 拆分混合内容

不要因为一个章节标题含"配置""策略"就把整个章节标 `internal`。应逐卡判断：

- 事实观察（如"存储、光纤光缆已有公司披露亮眼中报"）→ `both`
- 投资动作（如"进入再配置区间""配置信号需重视"）→ `internal`
- 同一张卡中事实与建议混合时，拆为两张卡分别标注，或改写卡片文本使 `both` 版只保留事实描述。

目标是最大化 External 的信息丰富度，而非最小化合规风险。

### 表格标注

`dual` 模式下，表格与观点卡必须逐项标注 `audience`。External 会检查表题、表头和每一个单元格，未标 `both` 的表格不外发。

### 不确定时

任何不确定是否适合外发的内容都标为 `internal`，且不得在 `external_title`、`external_subtitle`、`share.external_*` 中放入敏感内容。

## 合规参考

判定外部版允许与禁止的内容时，阅读 `references/external-compliance.md`。
