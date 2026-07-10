---
name: document-structure-map
description: "Transform PDFs, DOCX files, Markdown notes, web pages, images, transcripts, reports, or other long documents into a single vertical structure-map long image/HTML: extract the document's argument skeleton, section hierarchy, key claims, evidence, comparison tables, matrices, ecosystem blocks, and render it as a shareable 文档总图/总截图/结构化长图/一图看懂. Use when the user asks to make a 总截图, 长图总结, 文档总图, 结构化图解, research map, one-page visual brief, or wants a document converted into a visual skeleton rather than a prose summary."
---

# Document Structure Map

## Purpose

Turn an arbitrary document into a vertical "structure map" long image/HTML. The default output should closely follow the provided reference screenshot: a white long canvas, a single black left-side timeline rail, pastel section labels placed left of the rail, dense pastel horizontal claim bars placed right of the rail, and white inset chart/table blocks connected to the same rail.

Default to Simplified Chinese output unless the user explicitly asks for another language.

## Workflow

1. Ingest the source document.
   - Prefer structured extraction for PDFs, DOCX, Markdown, HTML, spreadsheets, and slide decks.
   - Preserve original headings, tables, figures, dates, numbers, named entities, and source links.
   - For images or screenshots, OCR first, then manually sanity-check obvious recognition errors.
   - For finance, market, legal, medical, or current-events documents, verify time-sensitive facts with current sources when needed.

2. Build the argument skeleton before rendering.
   - Identify the document's core question, thesis, section logic, evidence chain, and unresolved risks.
   - Collapse repetition. A structure map is not a full abstract and not a paragraph-by-paragraph digest.
   - Use 4-8 major sections for most documents. Use 2-5 claim cards per section.
   - Use the section's left label as the visible section title, matching the reference image. Do not duplicate a large heading inside the right content column unless the document truly needs a subheading.
   - Make each card one sharp point: `bold label + one sentence + key evidence`.

3. Choose visual blocks deliberately.
   - Use `table` for direct comparisons across dimensions.
   - Use `matrix` for two-axis judgment, positioning, or prioritization.
   - Use `ecosystem` for value chains, modules, stakeholders, product stacks, or causal systems.
   - Use plain cards when the point is a sequential claim or conclusion.
   - Avoid decorative diagrams that do not add reasoning value.

4. Create a JSON map using `references/schema.md`.
   - Include all final text in the JSON. Do not leave placeholders.
   - Keep chart/table labels concise enough to fit in a narrow long-image layout.
   - Use exact numbers and dates from sources; do not round important figures silently.

5. Render with `scripts/render_structure_map.py`.
   - Example: `python scripts/render_structure_map.py input.json --output output.html`
   - Save user-facing deliverables under the current task's `outputs/` directory when available.
   - If the user needs an image, open the HTML in a browser at about 900-960px viewport width and export a full-page screenshot as PNG.

6. Verify the result.
   - Check the rendered page visually at desktop width and a narrow mobile width.
   - Ensure no text overlaps, section labels align with the left rail, tables fit, and long words wrap.
   - Confirm the map contains the document's actual logic, not generic summary prose.

## Writing Rules

- Write for a final reader who has not seen the conversation.
- Do not include process notes, extraction errors, or phrases like "根据文档" unless needed for attribution.
- Keep the main claim visible near the top.
- Prefer precise section titles: `产业需求演进`, `技术路径差异`, `竞争格局`, `风险与验证点`.
- Make claim labels stable and skimmable, for example `核心结论`, `证据`, `分歧`, `风险`, `下一步验证`.
- If the source contains contradictions, surface them as `分歧/待验证`, not as smoothed-over conclusions.

## Visual Standard

- Use a single white vertical canvas, not a carded web page.
- Use a 1-2px black/gray vertical rail at the left. Every claim bar and major visual block should have a thin horizontal connector back to the rail.
- Put section titles in compact pastel label blocks left of the rail. Use the exact role of the screenshot's colored labels: they anchor each section, not just decorative tags.
- Put content in dense pastel horizontal strips: light blue/purple for definitions, pale amber for demand/industry logic, pale green for NPO/landing logic, pale pink for competitive dynamics, pale cyan for technical route details, pale lavender for CPU packaging, pale orange for domestic substitution.
- Keep claim bars flat and compact: 2-4px radius, light tint, no heavy shadows, label in bold followed by one sentence.
- Render tables, matrices, and ecosystem blocks as white inset cards with light gray borders and subtle shadows, placed between claim strips when they clarify the section.
- Avoid large hero titles, decorative gradients, floating cards, oversized headings, and marketing-page composition.
- The output should look like a polished research long screenshot based on the reference image.

## Resources

- Read `references/schema.md` when preparing the JSON structure or deciding which block type to use.
- Read `references/reference-style.md` when the user asks to replicate the screenshot's colors, structure, or visual feel.
- Use `scripts/render_structure_map.py` to produce the HTML deliverable from the JSON map.
