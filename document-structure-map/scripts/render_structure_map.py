#!/usr/bin/env python3
"""Render a document structure-map JSON file into a reference-style HTML page."""

from __future__ import annotations

import argparse
import html
import json
from pathlib import Path
from typing import Any


PALETTE = {
    "blue": {"accent": "#4c56b8", "strip": "#f6f4ff", "label": "#f1efff"},
    "amber": {"accent": "#a66b16", "strip": "#fff7e8", "label": "#fff0d4"},
    "green": {"accent": "#20934f", "strip": "#ecfff0", "label": "#ddf7e3"},
    "pink": {"accent": "#c43d8b", "strip": "#fff0f8", "label": "#ffe3f3"},
    "cyan": {"accent": "#168f89", "strip": "#edfffc", "label": "#dff9f6"},
    "purple": {"accent": "#6255bf", "strip": "#f5f2ff", "label": "#eeeaff"},
    "orange": {"accent": "#b66f18", "strip": "#fff7ea", "label": "#fff0d9"},
    "gray": {"accent": "#5f6673", "strip": "#f7f8fb", "label": "#eef1f5"},
}

TONE_ORDER = ["blue", "amber", "green", "pink", "cyan", "purple", "orange", "gray"]


def esc(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def palette_for(tone: str | None, index: int) -> dict[str, str]:
    key = tone if tone in PALETTE else TONE_ORDER[index % len(TONE_ORDER)]
    return PALETTE[key]


def render_cards(cards: list[dict[str, Any]], pal: dict[str, str]) -> str:
    parts = []
    for card in cards:
        source = card.get("source")
        source_html = f'<div class="source">{esc(source)}</div>' if source else ""
        parts.append(
            f"""
            <article class="claim-card timeline-node" style="background:{pal['strip']}">
              <span class="claim-label">{esc(card.get("label", "Point"))}</span>
              <span class="claim-text">{esc(card.get("text", ""))}</span>
              {source_html}
            </article>
            """
        )
    return "\n".join(parts)


def render_table(block: dict[str, Any]) -> str:
    headers = block.get("headers", [])
    rows = block.get("rows", [])
    header_html = "".join(f"<th>{esc(header)}</th>" for header in headers)
    row_html = []
    for row in rows:
        row_html.append("<tr>" + "".join(f"<td>{esc(cell)}</td>" for cell in row) + "</tr>")
    return f"""
    <section class="visual-block table-block timeline-node">
      <h3>{esc(block.get("title", "Comparison"))}</h3>
      <div class="table-wrap">
        <table>
          <thead><tr>{header_html}</tr></thead>
          <tbody>{''.join(row_html)}</tbody>
        </table>
      </div>
    </section>
    """


def render_matrix(block: dict[str, Any]) -> str:
    points = []
    for point in block.get("points", []):
        pal = palette_for(point.get("tone"), 0)
        x = max(0, min(100, float(point.get("x", 50))))
        y = max(0, min(100, float(point.get("y", 50))))
        points.append(
            f"""
            <div class="matrix-point" style="left:{x}%; bottom:{y}%; background:{pal['accent']}">
              <span>{esc(point.get("label", ""))}</span>
            </div>
            """
        )
    return f"""
    <section class="visual-block matrix-block timeline-node">
      <h3>{esc(block.get("title", "Matrix"))}</h3>
      <div class="matrix">
        <div class="axis axis-x">{esc(block.get("x_axis", "X axis"))}</div>
        <div class="axis axis-y">{esc(block.get("y_axis", "Y axis"))}</div>
        <div class="quad-line vertical"></div>
        <div class="quad-line horizontal"></div>
        {''.join(points)}
      </div>
    </section>
    """


def render_ecosystem(block: dict[str, Any]) -> str:
    items = []
    for item in block.get("items", []):
        details = item.get("details", [])
        details_html = "".join(f"<li>{esc(detail)}</li>" for detail in details)
        items.append(
            f"""
            <article class="eco-item">
              <h4>{esc(item.get("title", "Module"))}</h4>
              <ul>{details_html}</ul>
            </article>
            """
        )
    return f"""
    <section class="visual-block ecosystem-block timeline-node">
      <h3>{esc(block.get("title", "Ecosystem"))}</h3>
      <div class="eco-grid">{''.join(items)}</div>
    </section>
    """


def render_note(block: dict[str, Any], pal: dict[str, str]) -> str:
    body = block.get("text") or block.get("body") or ""
    return f"""
    <section class="visual-block note-block timeline-node" style="background:{pal['strip']}">
      <h3>{esc(block.get("title", "Note"))}</h3>
      <p>{esc(body)}</p>
    </section>
    """


def render_blocks(blocks: list[dict[str, Any]], pal: dict[str, str]) -> str:
    rendered = []
    for block in blocks:
        kind = block.get("type", "note")
        if kind == "table":
            rendered.append(render_table(block))
        elif kind == "matrix":
            rendered.append(render_matrix(block))
        elif kind == "ecosystem":
            rendered.append(render_ecosystem(block))
        else:
            rendered.append(render_note(block, pal))
    return "\n".join(rendered)


def render_section(section: dict[str, Any], index: int) -> str:
    pal = palette_for(section.get("tone"), index)
    rail_label = section.get("tag") or section.get("title") or f"{index + 1:02d}"
    heading = section.get("heading", "")
    heading_html = f"<h2>{esc(heading)}</h2>" if heading else ""
    cards = render_cards(section.get("cards", []), pal)
    blocks = render_blocks(section.get("blocks", []), pal)
    return f"""
    <section class="map-section">
      <div class="section-label" style="background:{pal['label']}; color:{pal['accent']}">{esc(rail_label)}</div>
      <div class="section-body">
        {heading_html}
        <div class="cards">{cards}</div>
        {blocks}
      </div>
    </section>
    """


def render_hero(data: dict[str, Any], style: dict[str, Any]) -> str:
    if not style.get("show_title", False):
        return ""
    meta = data.get("meta", {})
    meta_bits = [meta.get("source"), meta.get("date"), meta.get("analyst")]
    meta_text = " | ".join(str(bit) for bit in meta_bits if bit)
    subtitle = data.get("subtitle", "")
    subtitle_html = f'<p class="subtitle">{esc(subtitle)}</p>' if subtitle else ""
    meta_html = f'<div class="meta">{esc(meta_text)}</div>' if meta_text else ""
    return f"""
    <header class="hero">
      <h1>{esc(data.get("title", "Document Structure Map"))}</h1>
      {subtitle_html}
      {meta_html}
    </header>
    """


def render_html(data: dict[str, Any]) -> str:
    style = data.get("style", {})
    sections = data.get("sections", [])
    section_html = "\n".join(render_section(section, index) for index, section in enumerate(sections))
    hero_html = render_hero(data, style)
    map_class = "map has-hero" if hero_html else "map"
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(data.get("title", "Document Structure Map"))}</title>
  <style>
    * {{ box-sizing: border-box; }}
    html, body {{ margin: 0; padding: 0; }}
    body {{
      background: #ffffff;
      color: #20232a;
      font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif;
      font-size: 14px;
      line-height: 1.56;
      letter-spacing: 0;
    }}
    .page {{
      width: min(960px, calc(100% - 28px));
      margin: 0 auto;
      padding: 28px 0 48px;
    }}
    .hero {{
      margin-left: 230px;
      margin-bottom: 20px;
      max-width: 700px;
    }}
    h1 {{
      margin: 0 0 8px;
      font-size: 24px;
      line-height: 1.22;
      letter-spacing: 0;
    }}
    .subtitle {{
      margin: 0;
      color: #596273;
      font-size: 14px;
    }}
    .meta {{
      margin-top: 8px;
      color: #8a92a0;
      font-size: 12px;
    }}
    .map {{
      position: relative;
      padding-left: 230px;
    }}
    .map::before {{
      content: "";
      position: absolute;
      left: 176px;
      top: 0;
      bottom: 0;
      width: 1.5px;
      background: #272727;
      border-radius: 2px;
    }}
    .map-section {{
      position: relative;
      margin: 0 0 34px;
      min-height: 40px;
    }}
    .section-label {{
      position: absolute;
      left: -230px;
      top: 0;
      width: 168px;
      min-height: 28px;
      padding: 6px 8px;
      border-radius: 2px;
      font-size: 14px;
      line-height: 1.18;
      font-weight: 800;
      overflow-wrap: anywhere;
    }}
    .section-label::after {{
      content: "";
      position: absolute;
      right: -8px;
      top: 15px;
      width: 8px;
      height: 1px;
      background: #272727;
    }}
    .section-body {{
      width: min(700px, 100%);
      padding-top: 0;
    }}
    h2 {{
      margin: 0 0 8px;
      color: #252b36;
      font-size: 15px;
      line-height: 1.28;
      letter-spacing: 0;
    }}
    .cards {{
      display: grid;
      gap: 8px;
    }}
    .timeline-node {{
      position: relative;
    }}
    .timeline-node::before {{
      content: "";
      position: absolute;
      left: -54px;
      top: 18px;
      width: 54px;
      height: 1px;
      background: #272727;
    }}
    .claim-card {{
      min-height: 34px;
      padding: 8px 12px;
      border-radius: 2px;
      box-shadow: none;
      overflow-wrap: anywhere;
    }}
    .claim-label {{
      font-weight: 800;
      color: #20232a;
    }}
    .claim-label::after {{
      content: "\\ff1a";
    }}
    .claim-text {{
      color: #262d3a;
    }}
    .source {{
      margin-top: 4px;
      color: #8a92a0;
      font-size: 11px;
    }}
    .visual-block {{
      margin-top: 14px;
      margin-bottom: 2px;
      background: #ffffff;
      border: 1px solid #e7e9ef;
      border-radius: 5px;
      padding: 16px;
      box-shadow: 0 2px 10px rgba(25, 35, 55, 0.06);
    }}
    .visual-block::before {{
      top: 20px;
    }}
    .visual-block h3 {{
      margin: 0 0 12px;
      color: #263241;
      font-size: 13px;
      line-height: 1.25;
      font-weight: 800;
    }}
    .table-wrap {{
      width: 100%;
      overflow-x: auto;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      font-size: 12px;
    }}
    th, td {{
      border: 1px solid #e0e4ea;
      padding: 8px 9px;
      text-align: left;
      vertical-align: middle;
      overflow-wrap: anywhere;
    }}
    th {{
      background: #f7f8fa;
      color: #242b36;
      font-weight: 800;
    }}
    .matrix {{
      position: relative;
      height: 280px;
      margin: 8px 8px 18px 42px;
      border-left: 1px solid #cbd1dc;
      border-bottom: 1px solid #cbd1dc;
      background:
        linear-gradient(#edf0f5 1px, transparent 1px),
        linear-gradient(90deg, #edf0f5 1px, transparent 1px);
      background-size: 50% 50%;
    }}
    .quad-line {{
      position: absolute;
      background: #d9dee7;
    }}
    .quad-line.vertical {{
      left: 50%;
      top: 0;
      width: 1px;
      height: 100%;
    }}
    .quad-line.horizontal {{
      left: 0;
      bottom: 50%;
      width: 100%;
      height: 1px;
    }}
    .axis {{
      position: absolute;
      color: #727b8a;
      font-size: 11px;
      font-weight: 700;
    }}
    .axis-x {{
      right: 0;
      bottom: -23px;
    }}
    .axis-y {{
      left: -36px;
      top: 0;
      writing-mode: vertical-rl;
      transform: rotate(180deg);
    }}
    .matrix-point {{
      position: absolute;
      width: 11px;
      height: 11px;
      border-radius: 50%;
      transform: translate(-50%, 50%);
      box-shadow: 0 0 0 4px rgba(255, 255, 255, 0.92);
    }}
    .matrix-point span {{
      position: absolute;
      left: 16px;
      top: -7px;
      min-width: 78px;
      color: #334155;
      font-size: 11px;
      font-weight: 800;
      white-space: normal;
    }}
    .eco-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(142px, 1fr));
      gap: 12px;
    }}
    .eco-item {{
      min-height: 92px;
      border: 1px solid #e4e8ef;
      border-radius: 3px;
      padding: 11px 12px;
      background: #ffffff;
    }}
    .eco-item h4 {{
      margin: 0 0 7px;
      color: #263241;
      font-size: 13px;
      line-height: 1.25;
    }}
    .eco-item ul {{
      margin: 0;
      padding-left: 17px;
      color: #4b5563;
      font-size: 12px;
    }}
    .note-block p {{
      margin: 0;
      color: #374151;
      font-size: 13px;
    }}
    @media (max-width: 760px) {{
      .page {{
        width: calc(100% - 18px);
        padding-top: 18px;
      }}
      .hero {{
        margin-left: 112px;
      }}
      .map {{
        padding-left: 112px;
      }}
      .map::before {{
        left: 88px;
      }}
      .section-label {{
        left: -112px;
        width: 82px;
        font-size: 11px;
        padding: 5px 6px;
      }}
      .section-label::after {{
        right: -6px;
        width: 6px;
      }}
      .timeline-node::before {{
        left: -24px;
        width: 24px;
      }}
      .section-body {{
        width: 100%;
      }}
      .claim-card {{
        padding: 7px 9px;
      }}
      .visual-block {{
        padding: 11px;
      }}
    }}
  </style>
</head>
<body>
  <main class="page">
    {hero_html}
    <div class="{map_class}">
      {section_html}
    </div>
  </main>
</body>
</html>
"""


def load_json(path: str) -> dict[str, Any]:
    if path == "-":
        import sys

        return json.load(sys.stdin)
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def main() -> None:
    parser = argparse.ArgumentParser(description="Render document structure-map JSON to HTML.")
    parser.add_argument("input", help="Input JSON file, or '-' for stdin.")
    parser.add_argument("--output", "-o", help="Output HTML file. Defaults to input basename with .html.")
    args = parser.parse_args()

    data = load_json(args.input)
    if args.output:
        output = Path(args.output)
    elif args.input == "-":
        output = Path("structure-map.html")
    else:
        output = Path(args.input).with_suffix(".html")

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(render_html(data), encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
