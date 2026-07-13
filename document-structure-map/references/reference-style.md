# Reference Screenshot Style

Use this visual specification when the user asks to replicate the provided reference image, make a "总截图", or preserve its color and structure.

## Structure

- Canvas: tall white page, compact margins, no hero section by default.
- Main rail: one thin black/dark-gray vertical line running down the left side.
- Section label: pastel colored rectangle placed directly over the rail. The rail is visually interrupted by the label background, then continues below it; the label contains the section title and serves as the main visible heading.
- Connector lines: every claim strip and major visual block should connect back to the rail with a thin horizontal black/dark-gray line.
- Content column: right of the rail, medium width, high information density.
- Visual blocks: white inset cards inside the content column for tables, matrices, charts, and ecosystem grids.
- Screenshot/export viewport: use about 900-960px width for the final full-page screenshot so the left rail, labels, and content column match the reference density.

## Color Sequence

Use the sequence below unless the content has a stronger semantic reason:

| Use case | tone | Label tint | Content tint | Accent |
|---|---|---:|---:|---:|
| Core definitions / concepts | blue | `#f1efff` | `#f6f4ff` | `#4c56b8` |
| Industry demand / commercial logic | amber | `#fff0d4` | `#fff7e8` | `#a66b16` |
| NPO technical / landing logic | green | `#ddf7e3` | `#ecfff0` | `#20934f` |
| Domestic or competitive dynamics | pink | `#ffe3f3` | `#fff0f8` | `#c43d8b` |
| Technical route / CPO details | cyan | `#dff9f6` | `#edfffc` | `#168f89` |
| CPU packaging / adjacent technology | purple | `#eeeaff` | `#f5f2ff` | `#6255bf` |
| Localization / substitution / customers | orange | `#fff0d9` | `#fff7ea` | `#b66f18` |

## Typography

- Section labels: bold, compact, 14-15px equivalent, no long wrapping if avoidable.
- Claim strips: bold label plus concise text, 13-14px equivalent.
- Tables and diagrams: smaller but readable, 12-13px equivalent.
- Use dense spacing. The reference image prioritizes scan efficiency over spacious editorial layout.

## Content Rules

- Put the full section title in `section.tag` for the rail label.
- Keep `section.heading` empty unless a right-column subheading is necessary.
- Write 2-5 claim cards per section.
- Use a table only for direct comparisons.
- Use a matrix only for two-axis positioning or maturity/uncertainty tradeoffs.
- Use an ecosystem block only for supply chains, module systems, or stakeholder maps.
- Do not add decorative icons unless they encode category meaning inside an ecosystem block.
