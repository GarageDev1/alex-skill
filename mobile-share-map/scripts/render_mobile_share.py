#!/usr/bin/env python3
"""Render compliant internal/external mobile-share long images from one JSON map."""

from __future__ import annotations

import argparse
import copy
import html
import json
import re
from pathlib import Path
from typing import Any


SENSITIVE_PATTERNS = (
    r"评级|买入|卖出|持有|增持|减持|推荐|推票|建仓|交易指令|"
    r"估值|市值|目标价|目标市值|股价|发行价|P/E|P/B|P/S|"
    r"(?<![A-Za-z])(?:IPO|PE|PB|PS|DCF|IRR)(?![A-Za-z])|"
    r"收益率|回报率|上行空间|下行空间|安全边际|涨幅|跌幅|低估|高估|便宜|贵|"
    r"Bull|Base|Bear|概率加权|隐含价|隐含市值|"
    r"盈利预测|营收预测|利润预测|EPS 预测|预计.*(营收|收入|利润|出货)|预测.*(营收|收入|利润|出货)|"
    r"20\d{2}[EF]"
)
SENSITIVE_RE = re.compile(SENSITIVE_PATTERNS, flags=re.IGNORECASE)

# `single` is allowed for neutral industry/company information, including capacity forecasts,
# IPO process facts, and technical terms such as "base die". Only explicit securities-investment
# language should force the dual-version workflow. External filtering remains intentionally broader.
SINGLE_INVESTMENT_PATTERNS = (
    r"投资建议|评级|买入|卖出|持有|增持|减持|推票|建仓|加仓|减仓|仓位|止损|止盈|交易指令|"
    r"估值|目标价|目标市值|股价|发行价|P/E|P/B|P/S|"
    r"(?<![A-Za-z])(?:PE|PB|PS|DCF|IRR)(?![A-Za-z])|"
    r"收益率|回报率|上行空间|下行空间|安全边际|隐含价|隐含市值|"
    r"(?:Bull|Base|Bear)\s*(?:case|情景)|"
    r"(?:股价|估值|市值|标的).{0,8}(?:低估|高估)|(?:低估|高估).{0,8}(?:股价|估值|市值|标的)"
)
SINGLE_INVESTMENT_RE = re.compile(SINGLE_INVESTMENT_PATTERNS, flags=re.IGNORECASE)

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
TONE_ORDER = list(PALETTE)
SKILL_ROOT = Path(__file__).resolve().parents[1]
BRAND_CONFIG_PATH = SKILL_ROOT / "assets" / "brand.json"


def esc(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def clamp_int(value: Any, default: int, low: int, high: int) -> int:
    try:
        return max(low, min(high, int(value)))
    except (TypeError, ValueError):
        return default


def safe_color(value: Any, default: str = "#e45b3f") -> str:
    text = str(value or "").strip()
    if len(text) in (4, 7) and text.startswith("#") and all(
        char in "0123456789abcdefABCDEF" for char in text[1:]
    ):
        return text
    return default


def contains_sensitive(value: Any) -> bool:
    return bool(SENSITIVE_RE.search(str(value or "")))


def contains_investment_advice(value: Any) -> bool:
    return bool(SINGLE_INVESTMENT_RE.search(str(value or "")))


def load_brand_config() -> dict[str, Any]:
    """Load the editable skill-level brand settings; fall back safely if unavailable."""
    try:
        value = json.loads(BRAND_CONFIG_PATH.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}
    return value if isinstance(value, dict) else {}


def asset_uri(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if re.match(r"^(?:https?:|data:|file:)", text, flags=re.IGNORECASE):
        return text
    path = Path(text)
    if not path.is_absolute():
        path = SKILL_ROOT / path
    return path.resolve().as_uri() if path.exists() else ""


def brand_settings(share: dict[str, Any]) -> dict[str, Any]:
    """Prefer the skill-level brand file; allow an input JSON to opt into an override."""
    brand = load_brand_config()
    use_input_override = share.get("override_brand") is True

    def value(key: str, fallback: Any = "") -> Any:
        if use_input_override and key in share:
            return share[key]
        if key in brand:
            return brand[key]
        if key in share:
            return share[key]
        return fallback

    points = value("footer_points", [])
    return {
        "masthead_text": value("masthead_text", "智富界 · 用AI炒股的人都在这"),
        "footer_title": value("footer_title", "完整研报加入智富界交流群"),
        "footer_intro": value("footer_intro", ""),
        "footer_points": points if isinstance(points, list) else [],
        "footer_background": asset_uri(value("footer_background", "")),
        "qr_image": asset_uri(value("qr_image", "")),
        "qr_label": value("qr_label", "扫码获取完整研报"),
    }


def palette_for(tone: str | None, index: int, previous_tone: str | None = None) -> tuple[str, dict[str, str]]:
    """Select a tone while preventing two adjacent sections from sharing a color."""
    selected_tone = tone if tone in PALETTE else TONE_ORDER[index % len(TONE_ORDER)]
    if selected_tone == previous_tone:
        selected_tone = TONE_ORDER[(TONE_ORDER.index(selected_tone) + 1) % len(TONE_ORDER)]
    return selected_tone, PALETTE[selected_tone]


def is_external_safe(node: dict[str, Any], fields: tuple[str, ...]) -> bool:
    if node.get("audience") != "both":
        return False
    return not any(contains_sensitive(node.get(field)) for field in fields)


def block_values(block: dict[str, Any]) -> list[str]:
    """Return every visible string in a visual block for compliance checks."""
    values = [str(block.get("title", "")), str(block.get("text", ""))]
    values.extend(str(value) for value in block.get("headers", []))
    values.extend(str(value) for row in block.get("rows", []) for value in row)
    return values


def is_external_safe_block(block: dict[str, Any]) -> bool:
    return block.get("audience") == "both" and not any(contains_sensitive(value) for value in block_values(block))


def external_variant(data: dict[str, Any]) -> dict[str, Any]:
    """Apply an allowlist first, then remove any residual sensitive content."""
    result = copy.deepcopy(data)
    title = result.get("external_title") or result.get("title") or "公司业务研究要点"
    subtitle = result.get("external_subtitle") or result.get("subtitle") or "聚焦业务、产品、技术与公开历史事实。"
    result["title"] = title if not contains_sensitive(title) else "公司业务研究要点"
    result["subtitle"] = subtitle if not contains_sensitive(subtitle) else "聚焦业务、产品、技术与公开历史事实。"
    result["meta"] = {"source": "公开资料整理", "date": result.get("meta", {}).get("date", "")}
    share = result.setdefault("share", {})
    safe_share_defaults = {
        "eyebrow": "智富界 · 公司研究",
        "badge": "",
        "footer_title": "完整研报加入智富界交流群",
        "external_footer_text": "获取更多公司与产业研究",
        "qr_label": "扫码获取完整研报",
    }
    for key, default in safe_share_defaults.items():
        if contains_sensitive(share.get(key)):
            share[key] = default

    external_sections = []
    for section in result.get("sections", []):
        if not is_external_safe(section, ("title", "tag")):
            continue
        cards = []
        for card in section.get("cards", []):
            if is_external_safe(card, ("label", "text")):
                cleaned = {"label": card.get("label", "要点"), "text": card.get("text", "")}
                cards.append(cleaned)
        blocks = []
        for block in section.get("blocks", []):
            if is_external_safe_block(block):
                blocks.append(
                    {
                        "type": block.get("type", "table"),
                        "title": block.get("title", ""),
                        "text": block.get("text", ""),
                        "headers": block.get("headers", []),
                        "rows": block.get("rows", []),
                    }
                )
        if cards or blocks:
            external_sections.append(
                {
                    "title": section.get("title", "业务研究要点"),
                    "tag": section.get("tag") or section.get("title", "业务研究要点"),
                    "tone": section.get("tone"),
                    "cards": cards,
                    "blocks": blocks,
                }
            )
    result["sections"] = external_sections
    if not external_sections:
        raise ValueError("External 没有可发布内容：请将允许公开展示的章节和观点卡显式标为 audience: both。")
    brand = brand_settings(result["share"])
    visible_values = [result["title"], result["subtitle"], result["meta"].get("date", "")]
    visible_values.extend(str(result["share"].get(key, "")) for key in safe_share_defaults)
    visible_values.extend(
        [brand["masthead_text"], brand["footer_title"], brand["footer_intro"], brand["qr_label"]]
    )
    visible_values.extend(
        str(point.get("label", "")) + str(point.get("text", ""))
        for point in brand["footer_points"]
        if isinstance(point, dict)
    )
    for section in external_sections:
        visible_values.extend((section["title"], section.get("tag", "")))
        visible_values.extend(value for card in section["cards"] for value in (card["label"], card["text"]))
        visible_values.extend(value for block in section["blocks"] for value in block_values(block))
    if any(contains_sensitive(value) for value in visible_values):
        raise ValueError("External 仍含敏感投资内容，请修正 JSON 的外部文案或 audience 标记。")
    return result


def internal_variant(data: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(data)
    for section in result.get("sections", []):
        section.pop("audience", None)
        for card in section.get("cards", []):
            card.pop("audience", None)
        for block in section.get("blocks", []):
            block.pop("audience", None)
    return result


def visible_content_values(data: dict[str, Any]) -> list[str]:
    """Collect strings that can appear in a rendered single-version image."""
    values = [str(data.get("title", "")), str(data.get("subtitle", ""))]
    meta = data.get("meta", {})
    values.extend(str(meta.get(key, "")) for key in ("source", "date", "analyst"))
    share = data.get("share", {})
    brand = brand_settings(share)
    values.extend(
        str(value)
        for value in (
            brand["masthead_text"],
            brand["footer_title"],
            brand["footer_intro"],
            brand["qr_label"],
        )
    )
    values.extend(
        str(point.get("label", "")) + str(point.get("text", ""))
        for point in brand["footer_points"]
        if isinstance(point, dict)
    )
    for section in data.get("sections", []):
        values.extend((str(section.get("title", "")), str(section.get("tag", ""))))
        values.extend(
            str(value)
            for card in section.get("cards", [])
            for value in (card.get("label", ""), card.get("text", ""))
        )
        values.extend(
            str(value)
            for block in section.get("blocks", [])
            for value in block_values(block)
        )
    return values


def single_variant(data: dict[str, Any]) -> dict[str, Any]:
    """Prepare one neutral research image and reject investment-sensitive content."""
    result = internal_variant(data)
    if any(contains_investment_advice(value) for value in visible_content_values(result)):
        raise ValueError(
            "single 模式检测到评级、估值、交易建议或其他投资敏感内容；"
            "请改用 output_policy: dual，或先向用户确认输出策略。"
        )
    return result


def render_footer(data: dict[str, Any], mode: str) -> str:
    share = data.get("share", {})
    brand = brand_settings(share)
    title = brand["footer_title"]
    intro = brand["footer_intro"]
    background = (
        f'<img class="footer-background" src="{esc(brand["footer_background"])}" alt="">'
        if brand["footer_background"]
        else ""
    )
    points = "".join(
        f'<div class="footer-point"><b>{esc(point.get("label", ""))}</b>{esc(point.get("text", ""))}</div>'
        for point in brand["footer_points"]
        if isinstance(point, dict)
    )
    if brand["qr_image"]:
        qr_html = f'<img class="qr-image" src="{esc(brand["qr_image"])}" alt="{esc(brand["qr_label"])}">'
    else:
        qr_html = '<div class="qr-placeholder"><span>二维码</span><small>预留位</small></div>'
    return f"""
    <footer class="mobile-footer">
      {background}
      <div class="footer-copy"><div class="footer-title">{esc(title)}</div><div class="footer-intro">{esc(intro)}</div><div class="footer-points">{points}</div></div>
      <div class="footer-qr">{qr_html}<div class="qr-label">{esc(brand["qr_label"])}</div></div>
    </footer>
    """


def render_sections(sections: list[dict[str, Any]]) -> str:
    rendered = []
    previous_tone = None
    for index, section in enumerate(sections):
        selected_tone, palette = palette_for(section.get("tone"), index, previous_tone)
        previous_tone = selected_tone
        cards = "".join(
            f'<article class="claim-card timeline-node" style="background:{palette["strip"]}"><span class="claim-label">{esc(card.get("label", "要点"))}：</span>'
            f'<span>{esc(card.get("text", ""))}</span></article>'
            for card in section.get("cards", [])
        )
        blocks = []
        for block in section.get("blocks", []):
            if block.get("type", "table") != "table":
                continue
            headers = "".join(f"<th>{esc(value)}</th>" for value in block.get("headers", []))
            rows = "".join(
                "<tr>" + "".join(f"<td>{esc(value)}</td>" for value in row) + "</tr>"
                for row in block.get("rows", [])
            )
            if headers and rows:
                title = f'<h3>{esc(block.get("title", "数据对照"))}</h3>' if block.get("title") else ""
                blocks.append(
                    f'<section class="visual-block timeline-node">{title}<div class="table-wrap"><table>'
                    f'<thead><tr>{headers}</tr></thead><tbody>{rows}</tbody></table></div></section>'
                )
        tag = section.get("tag") or section.get("title", "研究要点")
        rendered.append(
            f'<section class="map-section"><div class="section-label" '
            f'style="background:{palette["label"]};color:{palette["accent"]}">{esc(tag)}</div>'
            f'<div class="section-body">{cards}{"".join(blocks)}</div></section>'
        )
    return "\n".join(rendered)


def render_html(data: dict[str, Any], mode: str) -> str:
    style = data.get("style", {})
    share = data.get("share", {})
    safe_top = clamp_int(style.get("safe_top_px"), 132, 72, 240)
    accent = safe_color(style.get("accent"))
    badge = data.get("meta", {}).get("date", "") if mode == "external" else (share.get("badge") or data.get("meta", {}).get("date", ""))
    brand = brand_settings(share)
    eyebrow = brand["masthead_text"]
    header_qr = (
        f'<div class="header-qr"><img class="header-qr-image" src="{esc(brand["qr_image"])}" alt="{esc(brand["qr_label"])}">'
        f'<div class="header-qr-label">{esc(brand["qr_label"])}</div></div>'
        if brand["qr_image"]
        else ""
    )
    mode_label = "内部研究版" if mode == "internal" else ""
    compliance_html = f'<div class="compliance">{esc(mode_label)}</div>' if mode_label else ""
    return f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(data.get("title", "手机转发研究长图"))}</title><style>
*{{box-sizing:border-box}} html,body{{margin:0;padding:0;background:#eef1f5}} body{{font-family:"Microsoft YaHei","PingFang SC","Noto Sans CJK SC",Arial,sans-serif;color:#17191f}}
.page{{width:1080px;margin:0 auto;padding:0 72px;background:#fff;overflow:hidden;--accent:{accent}}} .safe-area{{height:{safe_top}px;margin:0 -72px;background:#fff}}
.masthead{{display:flex;align-items:center;justify-content:space-between;min-height:194px;padding:0 0 18px;border-bottom:3px solid #191b20;font-size:25px;font-weight:800;letter-spacing:3px}} .masthead-right{{display:flex;align-items:center;gap:24px}} .badge{{color:var(--accent);letter-spacing:1px}} .header-qr{{display:grid;justify-items:center;gap:6px;padding:8px 8px 7px;border:1px solid #d7dce5;background:#f7f9fc;font-size:21px;line-height:1.2;letter-spacing:0;color:#455066}} .header-qr-image{{width:142px;height:142px;padding:5px;border:1px solid #d7dce5;background:#fff;object-fit:contain}} .header-qr-label{{white-space:nowrap}}
.compliance{{display:inline-block;margin:28px 0 0;padding:7px 12px;background:#f3f5f8;color:#5f6673;font-size:21px;font-weight:800;letter-spacing:1px}} .hero{{max-width:936px;margin:30px 0 42px}} h1{{margin:0 0 16px;font-size:52px;line-height:1.18;letter-spacing:-2px}} .subtitle{{margin:0;color:#545966;font-size:27px;line-height:1.5}} .meta{{margin-top:16px;color:#8b909b;font-size:22px}}
.map{{position:relative;padding-left:112px}} .map::before{{content:"";position:absolute;left:64px;top:0;bottom:0;width:2px;background:#272727}} .map-section{{position:relative;margin:0 0 40px;min-height:44px}} .section-label{{position:relative;z-index:1;left:64px;display:table;width:auto;min-height:0;margin:0 0 18px -112px;padding:9px 11px;border-radius:2px;font-size:26px;line-height:1.25;font-weight:900;overflow-wrap:anywhere;transform:translateX(-50%)}} .section-label::after{{display:none}} .section-body{{width:824px}} .claim-card{{position:relative;margin:0 0 13px;padding:13px 16px;border-radius:3px;color:#353941;font-size:27px;line-height:1.56;overflow-wrap:anywhere}} .timeline-node::before{{content:"";position:absolute;left:-48px;top:31px;width:48px;height:2px;background:#272727}} .claim-label{{color:#20232a;font-weight:900}} .visual-block{{position:relative;margin:18px 0 4px;padding:15px 16px 17px;border:1px solid #dde2e9;border-radius:3px;background:#fff;color:#30343b;overflow:hidden}} .visual-block h3{{margin:0 0 11px;font-size:23px;line-height:1.3;color:#20232a}} .table-wrap{{width:100%;overflow:hidden}} table{{width:100%;border-collapse:collapse;table-layout:fixed}} th,td{{border:1px solid #dfe4eb;padding:9px 8px;font-size:19px;line-height:1.34;text-align:left;vertical-align:top;overflow-wrap:anywhere}} th{{background:#f1f4f8;color:#303640;font-weight:900}} td{{color:#4b515d}}
.mobile-footer{{position:relative;height:536px;margin:44px -72px 0;padding:0;overflow:hidden;color:#fff;background:#102c61}} .footer-background{{position:absolute;inset:0;width:100%;height:100%;object-fit:cover}} .footer-copy{{position:absolute;z-index:1;left:88px;top:198px;width:620px}} .footer-title{{font-size:35px;line-height:1.3;font-weight:900;text-shadow:0 0 18px rgba(63,188,255,.5)}} .footer-intro{{width:560px;margin-top:20px;color:rgba(255,255,255,.88);font-size:20px;line-height:1.75}} .footer-points{{display:grid;gap:5px;width:560px;margin-top:10px;color:rgba(255,255,255,.88);font-size:17px;line-height:1.35}} .footer-point b{{display:inline-block;width:25px;color:#9fdcff;font-size:20px}} .footer-qr{{position:absolute;z-index:1;left:776px;top:132px;width:206px;text-align:center}} .qr-image,.qr-placeholder{{display:flex;width:206px;height:206px;margin:0;background:#fff;object-fit:contain}} .qr-placeholder{{align-items:center;justify-content:center;flex-direction:column;border:10px solid #d9dee8;color:#657084;font-weight:900}} .qr-placeholder small{{font-size:18px;font-weight:500}} .qr-label{{margin-top:14px;font-size:22px;line-height:1.25;white-space:nowrap;text-shadow:0 0 12px rgba(28,120,255,.55)}}
</style></head><body><main class="page"><div class="safe-area"></div><div class="masthead"><span>{esc(eyebrow)}</span><div class="masthead-right"><span class="badge">{esc(badge)}</span>{header_qr}</div></div>{compliance_html}<header class="hero"><h1>{esc(data.get("title", "手机转发研究长图"))}</h1><p class="subtitle">{esc(data.get("subtitle", ""))}</p><div class="meta">{esc(data.get("meta", {}).get("source", ""))} | {esc(data.get("meta", {}).get("date", ""))}</div></header><div class="map">{render_sections(data.get("sections", []))}</div>{render_footer(data, mode)}</main></body></html>"""


def render_png(html_path: Path, png_path: Path) -> None:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError("缺少 Playwright，无法默认生成 PNG；请安装后重试，或仅为调试使用 --html-only。") from exc
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        page = browser.new_page(viewport={"width": 1080, "height": 1920}, device_scale_factor=1)
        page.goto(html_path.resolve().as_uri(), wait_until="networkidle")
        if page.evaluate("document.documentElement.scrollWidth") != 1080:
            raise RuntimeError("渲染出现水平溢出，拒绝输出图片。")
        page.locator(".page").screenshot(path=str(png_path))
        browser.close()


def write_mode(data: dict[str, Any], mode: str, output_dir: Path, stem: str, html_only: bool) -> list[Path]:
    if mode == "single":
        variant = single_variant(data)
        html_path = output_dir / f"{stem}.html"
    else:
        variant = internal_variant(data) if mode == "internal" else external_variant(data)
        html_path = output_dir / f"{stem}-{mode}.html"
    html_path.write_text(render_html(variant, mode), encoding="utf-8")
    outputs = [html_path]
    if not html_only:
        png_path = output_dir / (f"{stem}.png" if mode == "single" else f"{stem}-{mode}.png")
        render_png(html_path, png_path)
        outputs.append(png_path)
    return outputs


def main() -> None:
    parser = argparse.ArgumentParser(description="Render single or compliant dual mobile-share long images.")
    parser.add_argument("input", help="Input JSON file.")
    parser.add_argument("--output-dir", help="Output directory. Defaults to the input directory.")
    parser.add_argument(
        "--mode",
        choices=("auto", "single", "both", "internal", "external"),
        default="auto",
        help="Default: auto. Reads output_policy=single|dual from the input JSON.",
    )
    parser.add_argument("--html-only", action="store_true", help="Debug only: skip PNG output.")
    args = parser.parse_args()
    input_path = Path(args.input)
    data = json.loads(input_path.read_text(encoding="utf-8"))
    output_dir = Path(args.output_dir) if args.output_dir else input_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    if args.mode == "auto":
        policy = data.get("output_policy")
        if policy == "single":
            modes = ("single",)
        elif policy == "dual":
            modes = ("internal", "external")
        else:
            raise ValueError(
                "缺少有效的 output_policy。请先判断内容属于中性信息整理（single）"
                "还是含具体标的投资建议（dual）；无法判断时先询问用户。"
            )
    else:
        modes = ("internal", "external") if args.mode == "both" else (args.mode,)
    for mode in modes:
        for output in write_mode(data, mode, output_dir, input_path.stem, args.html_only):
            print(output)


if __name__ == "__main__":
    main()
