---
name: xhs-carousel-renderer
description: Render Chinese articles, reports, explainers, case studies, and data-backed posts into a polished multi-page Xiaohongshu-style PNG carousel. Use when an agent needs to turn text or Markdown—including tables, local images, links, code, task lists, nested lists, footnotes, and rich emphasis—into coordinated 1080x1440 social cards with a title-only cover, automatic pagination, fixed branding, page numbers, and selectable classic, finance, editorial, or tech visual themes.
---

# 小红书轮播图渲染器

产出是完成的图片集，不是交互式网页。严格保留原始素材中的事实和数据。

## 工作流程

1. 阅读原始素材，确定标题、副标题、数据来源、核心数据、强调要点和输出目录。
2. 准备或修改渲染器输入时，阅读 [references/content-format.md](references/content-format.md)。选择或更改视觉主题时，阅读 [references/themes.md](references/themes.md)。**每次都必须**阅读 [references/narrative-style.md](references/narrative-style.md)——它规定了所有轮播图的写作语气、数字密度、钩子结构和组件用法。
3. 按照叙事风格指南改写原始素材。每页应以读者视角的钩子开头，数字应编进叙述（而非堆在数据面板里），整体读感应像聊天而非汇报。如果某个话题适合插入一段真实故事，可联网从高置信度来源搜索。将搜到的故事素材先写入一个草稿笔记文件（不得直接写入最终 Markdown），以便审计 sub-agent 稍后核查。
4. 编写一份 UTF-8 Markdown 输入文件，使用支持的 front matter、块级指令和行内标记。不得编造事实、数据、引用或品牌信息。
5. 验证输入：

   ```bash
   node "<skill-dir>/scripts/validate.mjs" <input.md>
   ```

6. 渲染轮播图：

   ```bash
   node "<skill-dir>/scripts/render.mjs" <input.md> --output <output-dir>
   ```

   该命令输出编号的 1080×1440 PNG 文件和 `manifest.json`。会根据内容自动生成所需数量的正文页，并默认添加封面（设置 `cover: false` 可跳过）。使用 `--theme classic|finance|editorial|tech` 可预览或覆盖 front matter 中的主题而无需编辑输入文件。

7. 检查第一张、一张有代表性的正文页和最后一张 PNG。确认无页面溢出、无块被裁切、页序连贯、重点强调与原意一致。
8. **审计（sub-agent）。** 启动一个独立的 sub-agent 对最终 Markdown 和渲染后的 PNG 做叙事风格基线审计：
   - 每页是否以读者视角的钩子开头？
   - 是否使用了僵硬的编号结构（一、二、三）作为主要组织方式？
   - 整体读感是否像聊天，而非研报？
   - 如果使用了联网搜索的故事素材，中间产物文件是否存在？关键事实是否经过交叉核查？
   审计发现问题时，修改 Markdown 并重新渲染后再交付。
9. 仅当自动分页产生了不理想的叙事断点时，才调整内容分组或插入 `:::pagebreak`。调整后重新渲染并复查。

## 内容决策

- 匹配已有轮播图或未指定主题时保留 `classic`。仅当主题符合内容或用户要求时才切换。
- 按 [references/narrative-style.md](references/narrative-style.md) 的叙事风格指南写作。每页应让读者感到在听一个故事，而非在读一份报告。
- 每页数字不超过约3个。需要展示数据时优先编入叙述。仅在三组以上可比数据确实需要并排对照时才使用 `:::metrics`。
- 多列对比使用 Markdown 表格，建议不超过5列、10行。
- 图片单独一行，使用 Markdown 图片语法。相对路径从输入文件所在目录解析；渲染器会将本地图片嵌入渲染负载。
- 代码块尽量控制在18行以内。长表格、图片、代码块和 callout 会作为整体保留在同一页，不跨页拆分。
- 使用 `:::marker` 标记一页中最重要的一个结论。
- `:::callout` 为可选组件。当收尾观点能为页面增色时使用；如果页面本身已经有强力的叙事收尾，可以跳过。渲染器使用 front matter 中的 `callout_label` 字段作为面板前缀（默认"AI观点"）；只需写观点本身。
- `{circle}` 仅用于短数字或短语。
- `{wavy}` 用于需要手绘波浪下划线的短结论。
- `{accent}` 或 `==...==` 少用。同一页超过三种强调样式会显得杂乱。
- 标题和后续解释要紧挨在一起。需要特定叙事节拍另起一页时使用手动分页，而非缩小字号。
- 文章标题只出现在封面。每个正文页以自己的 section 标题或内容块开头。
- 段落要精炼。渲染器按语义块分页，不会拆分表格、指标组、图片、代码块或 marker。
- 优先使用 Markdown 而非原始 HTML。原始 HTML 限于安全的布局标签；属性会被移除，可执行元素保持转义文本。

## 渲染要求

- 需要 Node.js 和 Playwright 及 Chromium 浏览器。渲染器会在当前工作区和 SKILL 的上级工作区搜索 `playwright`。
- 验证器或渲染器返回非零退出码时视为真实失败。修复输入或运行环境；不得谎称图片已生成。
- 交付 PNG 图片集时附上输出目录路径和页数。不要交付仅用于内部排版的生成 HTML。
