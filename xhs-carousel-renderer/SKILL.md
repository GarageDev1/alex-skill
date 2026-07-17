---
name: xhs-carousel-renderer
description: Render Chinese articles, reports, explainers, case studies, and data-backed posts into a polished multi-page Xiaohongshu-style PNG carousel. Use when an agent needs to turn text or Markdown—including tables, local images, links, code, task lists, nested lists, footnotes, and rich emphasis—into coordinated 1080x1440 social cards with a title-only cover, automatic pagination, fixed branding, page numbers, and selectable classic, finance, editorial, or tech visual themes.
---

# Xiaohongshu Carousel Renderer

Create finished image sets, not an interactive webpage. Preserve the source's facts and numbers exactly.

## Workflow

1. Read the source material and identify the title, optional subtitle, brand, data source, metrics, emphasized claims, and desired output directory.
2. Read [references/content-format.md](references/content-format.md) when preparing or modifying renderer input. Read [references/themes.md](references/themes.md) when selecting or changing the visual theme.
3. Write one UTF-8 Markdown input file using the supported front matter, block directives, and inline marks. Do not invent facts, metrics, citations, or brand claims.
4. Validate the input:

   ```bash
   node "<skill-dir>/scripts/validate.mjs" <input.md>
   ```

5. Render the carousel:

   ```bash
   node "<skill-dir>/scripts/render.mjs" <input.md> --output <output-dir>
   ```

   The command writes numbered 1080x1440 PNG files and `manifest.json`. It creates as many body pages as needed and adds a cover unless `cover: false` is set. Use `--theme classic|finance|editorial|tech` to preview or override the front-matter theme without editing the input.

6. Inspect the first, a representative body, and the last PNG. Confirm that no page overflows, no block is clipped, page order is coherent, and important emphasis matches the intended meaning.
7. Adjust content grouping or insert `:::pagebreak` only when automatic pagination produces a weak narrative break. Render again and recheck.

## Content decisions

- Preserve `classic` when matching an existing carousel or when no theme is requested. Choose another theme only when it fits the subject or the user asks for a different visual direction.
- Use `:::metrics` for comparable label/value data. Keep the values verbatim.
- Use a Markdown table for genuine multi-column comparisons. Prefer no more than 5 columns and 10 body rows per table.
- Use Markdown images on their own line for figures. Resolve relative paths from the input Markdown file; the renderer embeds local images into the render payload.
- Keep code blocks under 18 lines when possible. Long tables, images, code blocks, and callouts stay together rather than splitting across pages.
- Use `:::marker` for the single most important conclusion on a page.
- Use `:::callout` for a larger closing takeaway with a pale-yellow panel and gold left rule.
- Use `{circle}` for a short number or phrase only.
- Use `{wavy}` for a short claim that needs a hand-drawn underline.
- Use `{accent}` or `==...==` sparingly. More than three simultaneous emphasis styles makes a page noisy.
- Keep headings with the following explanation. Use manual page breaks rather than weakening font size when a specific narrative beat must start a new image.
- Put the article title on the cover only. Start every body page with its own section heading or content block.
- Prefer concise paragraphs. The renderer paginates semantic blocks and does not split tables, metric groups, images, code blocks, or marker callouts.
- Prefer Markdown over raw HTML. Raw HTML is restricted to safe layout tags; attributes are removed and executable elements remain escaped text.

## Rendering requirements

- Require Node.js and Playwright with a Chromium browser. The renderer searches the current workspace and the skill's parent workspace for `playwright`.
- Treat a nonzero validator or renderer exit as a real failure. Fix the input or runtime; do not claim images were generated.
- Deliver the PNG set together with the output directory path and page count. Do not deliver the generated HTML used internally for layout.
