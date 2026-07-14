# 样式 profile

版式与内容解耦:同一份 base point `.md`,选不同 profile 出不同版式的 docx。profile 只控制"封面 / 字体 / 页眉页脚 / 引用呈现"这类**渲染层**决策,**不动正文内容**。

`--profile` 传入,`md_to_docx.py` 据此分支。

---

## `default`(通用)

适用:笔记、wiki、综合研究、任何非研报文档。**主题色:深蓝 `1F4E79`**。

- 正文:宋体 / Times New Roman 双字体,10.5pt(五号),1.4 倍行距,段后 6pt,两端对齐。
- 标题:微软雅黑加粗;H1/H2 主题色、H1 带主题色底线,H3/H4 近黑;段前段后间距,不孤悬页底。
- 表格:三线表(顶/底主题色粗线 + 浅蓝表头底纹,无竖线),数字列自动右对齐。
- 引用:左侧主题色条 + 浅灰底纹 callout;📊/📂 图注走灰色小字 caption。
- 代码块:浅灰底纹 + 细边框,保留换行。
- 图表:ECharts→PNG 居中 6" 嵌入 + 自动「图 N」编号。
- 页眉(文档标题)+ 页脚(页码)。
- **无封面**,从第一个标题直接开始。

## `equity-research`(研报)

适用:带研报 frontmatter 的 Markdown。在 `default` 之上换**研报红主题色 `C00000`**(标题/三线表/callout/封面眉标统一吃红),加首次覆盖封面。

**封面 Page 1**(数据源 = frontmatter,**这是唯一"合成"而非"转换"的部分,要单独验**):
- `INITIATING COVERAGE` / 「首次覆盖」眉标
- 公司标题(`company` → `title`)与 ticker
- **rating box**:评级 / 现价 / 目标价 / 空间 / 市值 / 行业 / 分析师 / 日期

**当前字段映射**:

| rating box 项 | frontmatter 字段 | 缺失兜底 |
|---|---|---|
| 评级 | `rating`(BUY→买入) | 该行不显示 |
| 现价 | `current_price` | 该行不显示 |
| 目标价 | `target_price` | 该行不显示 |
| 空间 | `upside` | 该行不显示 |
| 市值 | `market_cap` | 该行不显示 |
| 行业 | `sector` | 该行不显示 |
| 分析师 | `analyst` | 该行不显示 |
| 日期 | `date` | 该行不显示 |

**当前边界**:52 周区间、EV、封面 thesis bullet、财务摘要表尚未实现,不得在交付说明中声称已经生成。终检以本文件和 `conversion-rules.md` 为准,不依赖外部 skill 的 style lock 或 checklist。

---

## 扩展新 profile

加一个 profile = 在 `md_to_docx.py` 的 profile 分支里加封面/字体/页脚逻辑,本文件补一节说明。正文转换逻辑(`render_blocks`)全 profile 共用,不要 fork。
