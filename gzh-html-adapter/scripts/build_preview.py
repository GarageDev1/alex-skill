#!/usr/bin/env python3
"""Build a local preview page around a validated WeChat HTML fragment."""

from __future__ import annotations

import argparse
from html import escape
from pathlib import Path

from validate_gzh_html import validate_html


CONTENT_MARKER = "<!-- GZH_CONTENT -->"
TITLE_MARKER = "{{TITLE}}"


def default_output_path(source: Path) -> Path:
    suffix = "_公众号正文"
    stem = source.stem
    if stem.endswith(suffix):
        stem = stem[: -len(suffix)]
    return source.with_name(f"{stem}_公众号预览.html")


def build_preview(source: Path, output: Path | None = None) -> Path:
    if not source.is_file():
        raise FileNotFoundError(f"找不到正文文件：{source}")

    content = source.read_text(encoding="utf-8").strip()
    if not content:
        raise ValueError("正文文件为空")
    validation = validate_html(content)
    if not validation.ok:
        details = "；".join(validation.errors)
        raise ValueError(f"正文校验未通过：{details}")

    template_path = Path(__file__).resolve().parent.parent / "assets" / "preview-template.html"
    template = template_path.read_text(encoding="utf-8")

    if template.count(CONTENT_MARKER) != 1 or template.count(TITLE_MARKER) != 1:
        raise ValueError("预览模板标记缺失或重复")

    destination = output or default_output_path(source)
    rendered = template.replace(TITLE_MARKER, escape(source.stem))
    rendered = rendered.replace(CONTENT_MARKER, content)
    destination.write_text(rendered, encoding="utf-8")
    return destination


def main() -> int:
    parser = argparse.ArgumentParser(
        description="为公众号正文片段生成带复制按钮的本地预览页"
    )
    parser.add_argument("source", type=Path, help="已校验的公众号正文 HTML")
    parser.add_argument("-o", "--output", type=Path, help="预览页输出路径")
    args = parser.parse_args()

    try:
        destination = build_preview(args.source, args.output)
    except (FileNotFoundError, OSError, ValueError) as exc:
        parser.error(str(exc))

    print(f"已生成公众号预览页：{destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
