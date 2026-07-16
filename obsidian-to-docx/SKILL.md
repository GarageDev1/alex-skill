---
name: obsidian-to-docx
description: 将 Obsidian Markdown 导出为排版规整、可打印的 Word(.docx)文档。支持 vault 中的笔记、研报和 wiki,包括 ECharts 代码块、wikilink 双链、Markdown 表格与 LaTeX。通过 Markdown 到 docx 的元素映射、Puppeteer 渲染 ECharts PNG 和中英文双字体完成转换;研报封面等版式由可插拔 style profile 提供。输入 `.md` 是内容来源,导出过程不重做研究、建模或写作。当用户说"导出成 Word""转成 docx""把这篇笔记/研报出个 Word 版""obsidian 导 word"时触发。
---

# Obsidian → docx 导出器

把一个或一组 Obsidian Markdown 渲染成排版规整的 `.docx`。Word 与 Obsidian、xuang.xyz 网页同为 base point 的输出端,内容以输入 `.md` 为准。本 skill 只做格式转换,不改写内容,不重做研究。

**默认语言/版式**:正文中文用宋体、英文/数字用 Times New Roman(研报 register);ECharts 块用 puppeteer 渲成 PNG 嵌入,与网页端像素级一致。

---

## 边界(这个 skill 做什么 / 不做什么)

- ✅ 做:markdown 元素 → docx 元素的映射、ECharts→PNG、wikilink 清洗、双字体、可插拔版式 profile。
- ❌ 不做:研究 / 建模 / 估值 / 写作。内容来自输入 `.md`,缺内容就回去补 `.md`,不在导出环节造内容。

> 输入可以来自任意上游或普通 vault 笔记。导出器不依赖其他 skill,只认 markdown。

---

## 转换能力(markdown 元素 → docx)

完整映射规则与边界见 [`references/conversion-rules.md`](./references/conversion-rules.md)。骨架已实现的核心:

| markdown | docx | 状态 |
|---|---|---|
| frontmatter | 剥离;profile=equity-research 时喂封面 | M1✓ / 封面 M3 |
| `#`~`####` 标题(含 ≤3 空格/tab 缩进、Setext `===`/`---`) | Heading 1-4,微软雅黑 + 主题色,H1 底线 | M1✓ / 美化 M2.5✓ |
| 段落 / `**粗**` / `*斜*` / `` `码` `` | 段落 + run 格式,1.4 行距两端对齐 | M1✓ / 美化 M2.5✓ |
| `-`/`1.` 列表 | List Bullet / List Number | M1✓ |
| markdown 表格 `\|` | 三线表 + 表头底纹 + 数据行统一左对齐 + 列宽按内容比例分配 | M1✓ / 美化 M2.5✓ |
| `>` 引用 / 图注(📊📂) | 引用→色条 callout;图注→灰字 caption | M1✓ / 美化 M2.5✓ |
| `[文字](url)` | 可点击超链接 | M1✓ |
| `[[页面\|显示名]]` | 取显示名,去双链语法 | M1✓ |
| `![[图片]]` / `![](路径)` 嵌入 | vault 内按文件名查图→居中嵌入;多图自动并排;缺失占位 | M1.5✓ |
| ` ```echarts ` 块 | puppeteer 渲 PNG 居中嵌入 | 渲染脚本✓ / 嵌入 M2 |
| `$...$` / `$$...$$` LaTeX | OMML 或图片 | M3 |

---

## 流水线 & 用法

```
输入 .md
  │  ① node render_echarts.js  抽 echarts 块 → puppeteer 渲 → chart_NN.png
  │  ② python md_to_docx.py    解析 md → python-docx 组装(--charts-dir 嵌入①的图)
  └→ 输出 .docx
```

**两步命令**(② 的 echarts 编号顺序与 ① 严格一致,都按文档出现顺序):

```bash
# ① 渲染图表(默认加载本 skill 的 assets/echarts.min.js,无需联网)
node scripts/render_echarts.js <input.md> <charts_dir>

# ② 组装 docx(本地图片自动嵌入;--vault-root 给图片兜底搜索根,默认=输入 md 所在目录)
python scripts/md_to_docx.py --input <input.md> --output <out.docx> \
       --charts-dir <charts_dir> [--vault-root <vault根>] [--profile default|equity-research]
```

**纯文字快速预览**(跳过图表,echarts 显示占位框):省略 `--charts-dir` 即可,无需 node/puppeteer。**本地图片不依赖 charts-dir,任何情况下都会嵌入**。

---

## 样式 profile(可插拔版式)

版式与 base point 解耦,见 [`references/style-profiles.md`](./references/style-profiles.md):

- **`default`** — 通用文档:双字体正文 + 规整标题/表格/图表。适合笔记、wiki、综合研究。
- **`equity-research`** — 研报:在 default 之上加研报红主题与首次覆盖封面。当前 rating box 支持评级/现价/目标价/空间/市值/行业/分析师/日期,评级译中文(BUY→买入);完整字段边界见 `references/style-profiles.md`。

---

## 可选增强:在 Word 末尾附口播稿

`oral-script` 不是本 skill 的依赖。仅当当前环境另行安装该 skill、且用户明确要求"附口播稿"时,才用它生成 5-10 分钟口播稿。把结果追加到输入文档的临时副本末尾,再导出该临时副本;不要修改源 `.md`。

- 口播稿**只进 Word 交付版,不写回源 `.md`**。
- 若环境没有 `oral-script`,跳过此增强并说明;核心 Markdown→DOCX 流程不受影响。
- 用户没要求时,默认不附口播稿。

---

## 依赖

从 skill 根目录执行可复现安装:

```bash
python -m pip install -r requirements.txt
npm ci
```

- Python 依赖由 `requirements.txt` 锁定:`python-docx`、`lxml`、`latex2mathml`。
- Node 依赖由 `package-lock.json` 锁定:`puppeteer-core`;要求 Node.js ≥22.12,浏览器复用系统 Chrome/Edge/Chromium。
- 浏览器按 `--chrome` → `CHROME_PATH` → Windows/macOS/Linux 常见路径 → `PATH` 探测。
- ECharts 默认加载 `assets/echarts.min.js`,无需 CDN;如确需覆盖,传 `--echarts <本地路径或URL>`。
- ECharts/Markdown 转换规范以本 skill 的 `references/conversion-rules.md` 为准,不依赖外部 Obsidian skill。

---

## 实现进度(骨架)

- **M1 ✓** 核心转换器:frontmatter/标题/段落/列表/表格/引用/inline/wikilink/双字体/echarts 占位 — 已测通(见 `examples/`)
- **解析修复(标题不再丢)**:`#` 标题支持行首 ≤3 个空格/tab 缩进(CommonMark 合规、与 Obsidian 渲染一致);新增 Setext 标题识别(`标题\n===`→H1、`标题\n---`→H2);未闭合代码块打 stderr 警告(防后续内容被当代码静默吞掉)。回归测试:`python tests/test_parse_blocks.py`(22 例)
- **表格列宽按内容分配**:列宽不再等宽——按每列内容视觉宽度(中文2/ascii1)定权重、min 随列数自适应、归一化到正文宽,内容多的列宽、内容少的列窄;gridCol + 每 cell tcW + fixed 布局,Word 严格尊重不被改写。回归测试:`python tests/test_tables.py`(7 不变量)
- **文档大标题自动补全**:default profile 下,文档若无 H1 则在正文顶部补一个大标题(取 frontmatter title → 输入 md 文件名,对齐 Obsidian 顶部把文件名显示为最大标题的行为);文档自带 H1 时不重复补,equity-research 走封面不补。`pick_title` 改用输入文件名兜底(原用输出文件名,输出名临时会让标题错)
- **M1.5 ✓** 本地图片嵌入:`![[name]]`(Obsidian embed,按文件名在 md目录 / `attachments/` / vault 递归查)+ `![](path / file:///)` → 居中嵌入;一个图块多图自动并排(1→大图 / 2→2列 / 3→3列 / 4→2×2 / 5+→3列,无边框表格)+ 间隔;找不到的图标灰色占位 — 已在「深港答辩」29 图密集文档端到端测通
- **M2 ✓** 接 `--charts-dir` 真图嵌入(居中 6" + 自动「图 N」编号)
- **M2.5 ✓** 样式美化层:主题色 profile(default 深蓝 / equity-research 红)、标题微软雅黑 + H1 底线 + keep_with_next、正文 1.4 行距两端对齐、三线表 + 数字右对齐、引用 callout / 图注 caption 分流、代码块底纹 + 换行修复、页眉页脚页码 — 两 profile 均经 Word→PDF 目检验收
- **M3 进行中** equity-research 封面扩展字段(52周/EV)、■ bullet、财务摘要表;LaTeX 已支持 Office OMML,不可用时保留原始公式文本。
