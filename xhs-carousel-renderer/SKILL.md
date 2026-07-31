---
name: xhs-carousel-renderer
description: Render Chinese articles, reports, explainers, case studies, and data-backed posts into a polished multi-page Xiaohongshu-style PNG carousel. Use when an agent needs to turn text or Markdown—including tables, local images, links, code, task lists, nested lists, footnotes, and rich emphasis—into coordinated 1080x1440 social cards with a title-only cover, automatic pagination, fixed branding, page numbers, and selectable classic, finance, editorial, or tech visual themes.
---

# 小红书轮播图渲染器

产出是完成的图片集，不是交互式网页。严格保留原始素材中的事实和数据。

## 编码

所有输入和输出文件使用 UTF-8 编码。在 Windows 环境下运行脚本时，确保终端编码为 UTF-8（`chcp 65001`），否则中文内容可能出现乱码。

## 工作流程

### 1. 确认信源

用户在要求生成轮播图之前，应提供信源文件。常见格式为 HTML、DOCX、PDF。

- 如果用户已经提供了信源，确认收到并记录格式和路径。
- 如果用户没有明确提供信源，**必须主动询问**，不得跳过。

后续所有内容以信源为准。除了叙事风格指南允许的联网搜索故事素材以外，不得引入信源之外的事实、数据或观点。

### 2. 环境检查

运行环境依赖检测脚本：

```bash
python "<skill-dir>/scripts/preflight.py"
```

脚本会检测 Node.js、Playwright 和浏览器。`[FOUND]` 表示确认存在，`[UNCERTAIN]` 表示无法确认——需要自行判断是否可以继续。需要 Node.js 和 Playwright 及 Chromium 浏览器；渲染器会在当前工作区和 SKILL 的上级工作区搜索 `playwright`。如果关键依赖缺失且无法解决，终止流程并告知用户。

### 3. 素材分析

阅读原始素材，确定以下要素：标题、副标题、数据来源、核心数据、强调要点。

准备或修改渲染器输入时，阅读 [references/content-format.md](references/content-format.md)。选择或更改视觉主题时，阅读 [references/themes.md](references/themes.md)。**每次都必须**阅读 [references/narrative-style.md](references/narrative-style.md)——它规定了所有轮播图的写作语气、数字密度、钩子结构和组件用法。

### 4. 改写与编写 Markdown

#### 改写

按照叙事风格指南改写原始素材。

- 每页应以读者视角的钩子开头。
- 数字应编进叙述（而非堆在数据面板里）。
- 整体读感应像聊天而非汇报。
- 如果某个话题适合插入一段真实故事，可联网从高置信度来源搜索。将搜到的故事素材先写入一个草稿笔记文件（不得直接写入最终 Markdown），以便审计 sub-agent 稍后核查。

编写一份 UTF-8 Markdown 输入文件，使用 [references/content-format.md](references/content-format.md) 定义的 front matter、块级指令和行内标记。不得编造事实、数据、引用或品牌信息。

#### 手机阅读适配

轮播图主要在手机上阅读，屏幕小、注意力短。

- **优先用表格而非图片**展示结构化数据。表格在小屏上可读性远优于缩小后的图表。
- **尽量少插入图片**。如果必须插入，使用简单、对比度高、元素少的图——复杂图表缩到手机上根本看不清。
- 图片不传达关键信息时，考虑用文字描述替代。

#### 结尾缩略图

正文内容结束后，使用 `:::thumbnails` 指令插入信源缩略图。渲染器会自动生成带标题、网格布局和引导文案的缩略图页面，不需要手动排版。只需提供图片：

```md
:::thumbnails
![](./source-page1.png)
![](./source-page2.png)
:::
```

缩略图提取规则：

- **PDF / DOCX 信源**：提取前2页作为缩略图图片文件。
- **网页信源**：截取网页关键区域的截图。
- **无法自动获取时**：询问用户提供缩略图，或跳过。

放1-4张缩略图（渲染器自动居中排成一行）。推荐2张。

#### 排版决策

- 文章标题只出现在封面。每个正文页以自己的 section 标题或内容块开头。
- 标题和后续解释要紧挨在一起。
- **默认不加 `:::pagebreak`**。渲染器的自动分页会把内容排满再换页，手动分页几乎总是造成页面底部大面积空白。只在自动分页产生了不可接受的叙事断点时，才插入极少量手动分页。写完初稿后先不加任何 pagebreak 渲染一次，确认哪些断点确实需要调整，再有针对性地插入。
- 段落要精炼。渲染器按语义块分页，不会拆分表格、指标组、图片、代码块或 marker。
- 同一页不超过三种强调样式，否则显得杂乱。
- 优先使用 Markdown 而非原始 HTML。

### 5. 验证格式

```bash
node "<skill-dir>/scripts/validate.mjs" <input.md>
```

验证器返回非零退出码时视为真实失败。修复输入后重新验证。

### 6. 风格自查

```bash
python "<skill-dir>/scripts/lint.py" <input.md>
```

脚本检测：正文"你"字、kicker 主观词（解读/深度分析/研判/点评）、僵硬中文编号结构、每 section 数字密度、正文 H1 标题、风险内容是否在 `:::risk` 内。脚本只报告发现，不做决策。根据报告修改 Markdown，修改后重新运行第5步和本步骤，直到满意为止。

### 7. 渲染

所有产物统一输出到工作区下的 `视频图/` 目录。目录不存在时自动创建。

```bash
mkdir -p "<workspace>/视频图"
node "<skill-dir>/scripts/render.mjs" <input.md> --output "<workspace>/视频图"
```

输出编号的 1080×1440 PNG 文件和 `manifest.json`。会根据内容自动生成所需数量的正文页，并默认添加封面（设置 `cover: false` 可跳过）。使用 `--theme classic|finance|editorial|tech` 可预览或覆盖 front matter 中的主题。渲染器返回非零退出码时视为真实失败。修复输入或运行环境；不得谎称图片已生成。

### 8. 审计

启动一个独立的 sub-agent 对最终 Markdown 和渲染后的 PNG 做叙事风格基线审计。

**文本审计**（读 Markdown）：

- 每页是否以读者视角的钩子开头？
- 是否使用了僵硬的编号结构（一、二、三）作为主要组织方式？
- 整体读感是否像聊天，而非研报？
- 如果使用了联网搜索的故事素材，中间产物文件是否存在？关键事实是否经过交叉核查？

**视觉审计**（看渲染后的 PNG，至少检查封面、一张正文页和最后一页）：

- 是否有页面溢出或内容被裁切？
- 页面底部是否有不合理的大面积空白（通常是手动 pagebreak 导致）？
- 强调样式（marker、accent、circle）渲染是否正确、与原意一致？
- 缩略图页的图片是否正常显示？

审计发现问题时，修改 Markdown 并重新渲染后再交付。仅当自动分页产生了不理想的叙事断点时，才调整内容分组或插入 `:::pagebreak`。

### 9. 最终检查与交付

检查第一张、一张有代表性的正文页和最后一张 PNG。确认：

- 无页面溢出
- 无块被裁切
- 页序连贯
- 重点强调与原意一致

交付 PNG 图片集时附上输出目录路径和页数。不要交付仅用于内部排版的生成 HTML。
