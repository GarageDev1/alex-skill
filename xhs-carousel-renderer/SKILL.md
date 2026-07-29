---
name: xhs-carousel-renderer
description: Render Chinese articles, reports, explainers, case studies, and data-backed posts into a polished multi-page Xiaohongshu-style PNG carousel. Use when an agent needs to turn text or Markdown—including tables, local images, links, code, task lists, nested lists, footnotes, and rich emphasis—into coordinated 1080x1440 social cards with a title-only cover, automatic pagination, fixed branding, page numbers, and selectable classic, finance, editorial, or tech visual themes.
---

# 小红书轮播图渲染器

产出是完成的图片集，不是交互式网页。严格保留原始素材中的事实和数据。

## 编码

所有输入和输出文件使用 UTF-8 编码。在 Windows 环境下运行脚本时，确保终端编码为 UTF-8（`chcp 65001`），否则中文内容可能出现乱码。

## 工作流程

按顺序执行以下步骤。每一步开始前，阅读对应的说明文件。

1. **确认信源** — 阅读 [steps/01-confirm-source.md](steps/01-confirm-source.md)
2. **环境检查** — 阅读 [steps/02-preflight.md](steps/02-preflight.md)
3. **素材分析** — 阅读 [steps/03-analyze.md](steps/03-analyze.md)
4. **改写与编写 Markdown** — 阅读 [steps/04-write.md](steps/04-write.md)
5. **验证格式** — 阅读 [steps/05-validate.md](steps/05-validate.md)
6. **风格自查** — 阅读 [steps/06-lint.md](steps/06-lint.md)
7. **渲染** — 阅读 [steps/07-render.md](steps/07-render.md)
8. **审计** — 阅读 [steps/08-audit.md](steps/08-audit.md)
9. **最终检查与交付** — 阅读 [steps/09-review.md](steps/09-review.md)
