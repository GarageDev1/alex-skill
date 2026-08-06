# 步骤 3a：JSON 合规审计（Sub-agent）

在 JSON 创建完成、渲染之前，启动一个 sub-agent 对 JSON 进行独立合规审查。

## Sub-agent 职责

将以下内容交给 sub-agent：
- 生成的 JSON 文件
- `references/external-compliance.md`
- `references/schema.md`

要求 sub-agent 检查以下维度并逐项列出发现：

1. **audience 标注完整性**（仅 dual 模式）：每个 section、card、block 是否都显式标注了 `audience`；缺失标注的条目列出位置。
2. **混合内容拆分**：是否存在 `audience: "both"` 的卡片中混入了投资动作语言；逐卡扫描，给出具体文本。
3. **External 表面安全性**：`external_title`、`external_subtitle`、所有 `audience: "both"` 的标题/正文/表格内容中是否包含 `external-compliance.md` 列出的禁止项。
4. **schema 合规**：`output_policy` 是否存在、`tone` 是否有效、表格列数是否超过 4 列、`meta.date` 和 `share.badge` 是否被手动写入。
5. **信息丰富度**：External 可见的章节和卡片数量是否过少（例如 dual 模式下 `both` 内容不足总内容的三分之一），提示可能遗漏了可公开的事实内容。
6. **AI生成痕迹**：扫描标题、副标题和全部卡片，检查以下问题：
   - 指令回声句式："不需要X，只需要Y""不是X，而是Y""X的本质是Y"
   - 设问词："为什么""如何""怎样""是否""吗""贵不贵""划不划算"
   - 破折号"——"连接前后半句（副标题应使用逗号）
   - 过度学术化术语：如"二阶导数""一阶导数"等可以用大白话替代的表述
   - 副标题后半句是否为大白话陈述句（一个普通人会对朋友说的话）

## 主 agent 处理方式

Sub-agent 的输出是建议清单。Main Agent 逐条审阅：

- **认同的**：修改 JSON 后继续。
- **不认同的**：忽略，无需解释原因。
- **不确定的**：向用户说明 sub-agent 的发现，由用户决定。

审阅完成并修改 JSON 后，继续步骤 3b（数据可靠性审计）。
