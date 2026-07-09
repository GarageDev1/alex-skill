# 结构化 JSON schema

`render.py` 吃的中间产物。一个顶层对象,`categories` 是主体。字段带 `?` 为可选。

## 顶层
| 字段 | 类型 | 说明 |
|---|---|---|
| `theme?` | string | `business-blue` / `magazine-light` / `tech-dark`。命令行 `--theme` 会覆盖它。默认 `business-blue`。 |
| `kicker?` | string | hero 顶部小字(领域/栏目),如 "AI 算力 · 光互联产业全景"。 |
| `title` | string | 主标题(思维导图中心主题)。 |
| `summary?` | string | hero 一段导语,一两句话交代全图讲什么。 |
| `tags?` | string[] | hero 底部的关键词标签,3–4 个最佳。 |
| `footer?` | string | 页脚一行小字。 |
| `categories` | object[] | 一级分支(分类),建议 4–8 个。 |

## category(一级分支)
| 字段 | 类型 | 说明 |
|---|---|---|
| `title` | string | 分类标题(彩色节点标题)。 |
| `color?` | string | 十六进制色,如 `#2f7fd4`。省略则按主题调色板按顺序自动分配(推荐省略,保证同主题内协调)。 |
| `items` | object[] | 该分类下的节点,建议 3–6 条。 |

## item 节点类型

### `leaf` — 文字要点(最常用)
```json
{"type":"leaf","term":"光模块功能定位","text":"作为数据传输载体,实现电信号与光信号的转换……"}
```
- `term`:加粗的引导词(可空;空则整条纯文字)。
- `text`:一句到两句的说明。写成人话,别套话。

### `table` — 对比表
```json
{"type":"table","caption":"三类光电互联技术核心属性对比",
 "columns":["技术类型","封装位置","集成度","信号损耗","功耗水平"],
 "widths":[20,28,null,null,null],
 "highlightRow":2,
 "rows":[
   {"cells":["传统外置光模块","服务器/交换机外部端口",{"meter":"l","label":"低"},{"meter":"h","label":"高"},{"meter":"h","label":"高"}]}
 ]}
```
- `columns`:表头。`widths?`:各列百分比宽,不定的给 `null`。
- `highlightRow?`:0 起的行号,给该行浅色底(用来点出"胜出/重点"行)。
- `rows[].cells`:每格是**字符串**或**等级条对象** `{"meter":"l|m|h","label":"低|中|高"}`。等级条用同色系三段格显示程度(`l`=1 格、`m`=2 格、`h`=3 格),替代红橙绿标签,避免配色杂乱。

### `chart-quadrant` — 象限散点图
```json
{"type":"chart-quadrant","caption":"AI 算力光互联技术方向分类矩阵",
 "xLabel":"场景需求确定性","yLabel":"当前技术市场渗透率",
 "points":[
   {"label":"NPO","x":0.72,"y":0.50,"emphasis":true},
   {"label":"CPO","x":0.83,"y":0.28,"emphasis":"soft"},
   {"label":"常规光模块","x":0.28,"y":0.80,"muted":true}
 ]}
```
- `xLabel` / `yLabel`:两轴含义(轴向都是"低→高")。
- `points[].x` / `.y`:0–1 归一化坐标(0=低,1=高)。
- `emphasis`:`true`=用分类主色实心强调点;`"soft"`=主色半透明(次强调);省略=普通。
- `muted`:`true`=灰色弱化点(背景/参照项)。

### `units` — 要素/硬件网格
```json
{"type":"units","caption":"NPO 产业链核心硬件体系",
 "units":[
   {"icon":"发","title":"光发送单元","items":["发光芯片"]},
   {"icon":"配","title":"外围配套部件","items":["高价值量光引擎","光纤连接器","外置光源"]}
 ]}
```
- 每个 `unit`:`icon`(一个字/符号,放主色圆角块里)、`title`、`items[]`(要点列表)。
- 自动按 unit 数排成 1–4 列(超过 4 个也封顶 4 列)。

## 约定
- 事实忠于输入,不杜撰、不外部取数;缺关键数据回问用户。
- 一条 `leaf` 别写太长(两三行封顶),过长就拆成两条或改用 `table`/`units`。
- 能结构化的信息(对比、分类、并列要素)优先用 `table`/`chart-quadrant`/`units`,别全堆成 `leaf`。
