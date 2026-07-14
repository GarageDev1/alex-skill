# markdown → docx 转换规则

本文件是 `md_to_docx.py` 的行为契约:每类 markdown 元素如何落到 docx，以及边界与已知坑。改脚本时同步改这里。

---

## 块级元素

| 元素 | 识别 | docx 落点 | 备注 |
|---|---|---|---|
| frontmatter | 文件开头 `---...---` | 剥离;`default` 丢弃,`equity-research` 喂封面 | 不进正文 |
| 标题 | `^#{1,6}\s+` | `add_heading(level=min(n,4))` + 微软雅黑加粗;H1/H2 主题色、H1 带 1pt 主题色底线,H3/H4 近黑;段前段后间距 + `keep_with_next`(不孤悬页底) | docx 内置 Heading 最多到 4,5/6 级降级;标题黑体/正文宋体是中文排版惯例 |
| 段落 | 连续非空行 | 合并成一段(软换行→空格) | 遇块级起始符断开 |
| 无序列表 | `^\s*[-*+]\s+` | `List Bullet` | 暂不支持多级缩进(M2+) |
| 有序列表 | `^\s*\d+\.\s+` | `List Number` | 同上 |
| 引用/图注 | `^>\s` | 全部行以 📊/📂 开头 → 图注 caption(9pt 灰小字,缩进);否则 → callout(左侧 3pt 主题色条 + 浅灰底纹,9.5pt) | base point 图注(📊看图要点/📂数据来源)走 caption 分支 |
| 表格 | 行含 `\|` 且下一行是 `---\|---` 分隔 | 三线表:顶/底 1.5pt 主题色线 + 表头底纹居中加粗,无竖线,行间 0.5pt 浅灰;**数字单元格自动右对齐**(`_NUM_CELL_RE`:货币符/千分位/%/倍数);**表头跨页每页重复**(`w:tblHeader`)+ 单行不跨页断开(`w:cantSplit`) | 单元格走 inline 解析,单元格段落收紧(行距 1.15 / 无段后距) |
| 图片 | 整行为 `![[name]]` / `![](path)`(可含前导 `-` 列表符) | 见下「图片嵌入」 | 连续图片行合并为一组并排 |
| 代码块 | ` ``` ` 包裹 | Consolas 9pt + 浅灰底纹(F5F5F5)+ 0.5pt 边框,**逐行 `add_break()` 保留换行** | 非 echarts 的普通代码;直接把 `\n` 塞进一个 run 会被 Word 折叠成一行,必须用 break |
| echarts 块 | ` ```echarts ` | 见下「ECharts」 | 单独处理 |

## inline 元素

正则单遍扫描(`INLINE_RE`),优先级 = 出现顺序:

| 语法 | docx | 备注 |
|---|---|---|
| `**粗**` | `run.bold` | |
| `*斜*` | `run.italic` | |
| `` `码` `` | Consolas 字体 | |
| `[文字](url)` | 可点击超链接(蓝 + 下划线) | `add_hyperlink` 建 external relationship |
| `[[页面\|显示名]]` | 纯文字 = 显示名(无管道则页面名) | **双链语法被清除**,docx 无双链概念 |

**已知边界**(骨架阶段可接受,M2+ 再硬化):
- inline 不支持嵌套(如 `**粗里带 *斜* **`)——按最外层匹配。
- 链接 URL 里若含 `)` 会截断——base point 来源 URL 极少出现,暂不处理。
- 表格不支持单元格内换行 / 合并单元格。

---

## 图片嵌入

Obsidian 笔记的图片有两种写法,本转换器都按**块级**处理(整行是图片才算,图文混排行不拆图):
- Obsidian embed `![[name.png]]`(可带 `|尺寸`,会被忽略)
- markdown image `![](path)` / `![](file:///...)` / `![](http...)`

**识别**(`IMAGE_LINE_RE`):一行 strip 后(允许前导 `-`/`*`/`+` 列表符)整行由一个或多个图片 embed 构成 → 图片块。`_is_block_start` 也认它,避免上一段把图片行吞进正文。**连续的图片行合并成一个图片块**(中间空行则分块)。

**文件查找**(`resolve_image`,按序命中即停):
1. Obsidian embed:`md目录/name` → `md目录/basename` → `md目录/attachments/basename` → `vault根/basename` → `vault根/attachments/basename` → 最后 `os.walk(vault根)` 全索引兜底(跳过 `.git`/`.obsidian`/`node_modules` 等)。
2. `![](file:///…)`:URL-decode 后取本地路径,存在才嵌。
3. `![](相对/绝对路径)`:相对 md 目录解析。
4. `![](http…)`:**本骨架不下载**,记为缺失。

`--vault-root` 控制兜底搜索根,默认 = 输入 md 所在目录(同目录 `attachments/` 命中率最高;跨目录图片传 vault 根)。

**并排布局**(`add_image_grid`,无边框居中表格,总宽 6.2"):一个图块内 1→单张大图、2→2列、3→3列、4→2×2、5+→每行 3 列;单图宽 = 总宽/列数(上限 5.5")。每组后补一个空段间隔。

**缺失/失败**:找不到的图 → 居中灰色 `［图片缺失:文件名］`;`add_picture` 抛错(非图片/损坏)→ `［无法嵌入:文件名］`。二者都不静默丢弃,便于交付前人工补图。

**已知边界**:图文混排行(同一行既有文字又有图)不拆分,整行若不满足"纯图片"判定就当普通段落(此时图片语法会残留为文本);多级列表里的图片仅识别单层前导列表符。

---

## ECharts(核心难点)

base point 的图表是 ` ```echarts ` + JSON,Word 无法 inline 渲染,必须先转 PNG。

1. **抽取**:`render_echarts.js` 正则抽所有 echarts 块,**按文档出现顺序**编号 `chart_01.png` ...
2. **渲染**:puppeteer-core + 系统 Chrome/Edge/Chromium 加载本 skill 的 `assets/echarts.min.js` → `setOption` → 截图。默认不访问网络;可用 `--echarts <本地路径或URL>` 显式覆盖。
   - JSON 顶层 `width`/`height`(900/506)用来定容器尺寸,**setOption 前 strip 掉**(ECharts 不认 root-level 尺寸),与 xuang.xyz `ResearchMarkdown.tsx` 逻辑一致。
   - `deviceScaleFactor: 2` 出 2× 高清图,适配打印。
3. **嵌入**:`md_to_docx.py` 的 echarts 计数与渲染编号**严格对齐**(同样按出现顺序),`--charts-dir/chart_NN.png` 存在则居中嵌入(6"宽)+ 图下自动补居中灰色「图 N」编号,否则灰色占位框。图与编号 `keep_with_next` 不跨页拆开。

**图名归属(关键约定)**:base point 的 echarts **不写 `title.text`**(H3 中文描述即图名,避免"一图两名")。所以:
- 渲出的 PNG **图上没有标题**。
- docx 端图名来自图前的 H3 文字;图下的 📊/📂 图注由 `>` 引用块承载。
- 二者都不能在转换中丢失。

**编号对齐的脆弱点**:两个脚本各自独立数 echarts 块。只要都「按文档顺序、JSON 解析失败也占位不跳号」就一致。`render_echarts.js` 对解析失败的块 push null 占位正是为此。

---

## 双字体(中英混排)

`set_run_fonts`:每个 run 设 `w:rFonts` 的 `ascii`/`hAnsi`=Times New Roman、`eastAsia`=宋体。Word 据此自动按字符选字体——英文数字 Times、中文宋体。`Normal` 样式也补 `eastAsia`,兜底未显式设置的文本。**标题例外**:中英文统一微软雅黑(`ZH_HEADING_FONT`),标题黑体 / 正文宋体。超链接 run 也显式带 rFonts + sz,避免落回 Word 默认 Calibri。

---

## 版式层(M2.5)

- **主题色**:`THEMES` 按 profile 切换——default 深蓝 `1F4E79` / equity-research 研报红 `C00000`,各配表头底纹色。标题、H1 底线、三线表上下线、callout 色条都吃主题色。
- **Normal 样式**:10.5pt + 1.4 倍行距 + 段后 6pt;正文段落两端对齐(JUSTIFY)。表格单元格/页眉脚/图片格里的段落用 `_tight()` 收紧(行距 1.15/1.0、无段后距),否则表格会被 Normal 的行距撑得虚胖。
- **页眉页脚**(`setup_header_footer`):页眉 = 文档标题(右对齐 8pt 灰 + 细底线;标题取 frontmatter `title` → 首个 H1 → 文件名),页脚 = 居中 PAGE 域页码。
- **列表**:段后距收到 3pt,列表项之间比正文段落更紧凑。

---

## LaTeX

`$...$` / `$$...$$` 使用 `latex2mathml` 转 MathML,再由 `lxml` 和 Microsoft Office 的 `MML2OMML.XSL` 转为可编辑 OMML。只有含 LaTeX 命令或上下标特征的 `$...$` 才视为公式,避免把货币误判为公式。

若 `latex2mathml`、XSLT 或 Office XSL 不可用,导出不中断,公式保留为带 `$`/`$$` 的原始文本。当前没有 PNG 公式 fallback。
