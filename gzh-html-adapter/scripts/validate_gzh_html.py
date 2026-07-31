#!/usr/bin/env python3
"""Validate a clean WeChat article fragment against this skill's rules."""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path


ALLOWED_TAGS = {
    "section",
    "p",
    "h1",
    "h2",
    "h3",
    "h4",
    "blockquote",
    "ul",
    "ol",
    "li",
    "table",
    "thead",
    "tbody",
    "tr",
    "th",
    "td",
    "span",
    "strong",
    "em",
    "a",
    "br",
    "hr",
}

VOID_TAGS = {"br", "hr"}
STYLE_REQUIRED_TAGS = ALLOWED_TAGS - {"span", "br"}
TABLE_ATTRIBUTES = {"style", "colspan", "rowspan"}
ATTRIBUTE_ALLOWLIST = {
    "span": {"leaf", "style"},
    "a": {"href", "style"},
    "th": TABLE_ATTRIBUTES,
    "td": TABLE_ATTRIBUTES,
}

FORBIDDEN_TAGS = {
    "img",
    "picture",
    "source",
    "svg",
    "canvas",
    "video",
    "audio",
    "iframe",
    "object",
    "embed",
    "style",
    "script",
    "link",
    "div",
    "button",
    "form",
    "input",
}

FORBIDDEN_STYLE_PATTERNS = (
    (re.compile(r"url\s*\(", re.I), "CSS url()"),
    (re.compile(r"data\s*:", re.I), "data URI"),
    (re.compile(r"background-image\s*:", re.I), "background-image"),
    (re.compile(r"list-style-image\s*:", re.I), "list-style-image"),
    (re.compile(r"position\s*:\s*(?:absolute|fixed|sticky)", re.I), "禁用定位"),
    (re.compile(r"float\s*:", re.I), "float"),
    (re.compile(r"display\s*:\s*grid", re.I), "CSS Grid"),
    (re.compile(r"var\s*\(\s*--", re.I), "CSS 变量"),
    (re.compile(r"@(?:media|import|keyframes|supports)", re.I), "CSS @ 规则"),
    (re.compile(r"(?:linear|radial|conic)-gradient\s*\(", re.I), "渐变"),
    (re.compile(r"(?:box|text)-shadow\s*:", re.I), "阴影"),
    (re.compile(r"filter\s*:", re.I), "滤镜"),
    (re.compile(r"#[0-9a-f]{3,8}\b", re.I), "十六进制颜色"),
    (re.compile(r"(?:rgb|hsl)a?\s*\(", re.I), "RGB/HSL 颜色"),
)

FORBIDDEN_DOCUMENT_PATTERNS = (
    (re.compile(r"<!doctype", re.I), "正文不能包含 DOCTYPE"),
    (re.compile(r"<\s*/?\s*(?:html|head|body)\b", re.I), "正文不能包含文档外壳"),
    (re.compile(r"\b(?:src|srcset)\s*=", re.I), "正文不能包含媒体源属性"),
)

ALLOWED_COLOR_VALUES = {"inherit", "currentcolor", "transparent"}
COLOR_PROPERTY = re.compile(
    r"^(?:background|background-color|(?:[a-z-]*-)?color)$",
    re.I,
)
BORDER_PROPERTY = re.compile(r"^border(?:-(?:top|right|bottom|left))?$", re.I)
BORDER_TOKEN = re.compile(
    r"^(?:"
    r"-?(?:\d+(?:\.\d+)?|\.\d+)(?:px|em|rem|%|pt)?"
    r"|none|solid|dashed|dotted|double|thin|medium|thick"
    r"|inherit|currentcolor|transparent"
    r")$",
    re.I,
)
HTTP_LINK = re.compile(r"^https?://", re.I)


@dataclass(frozen=True)
class ValidationResult:
    errors: tuple[str, ...]

    @property
    def ok(self) -> bool:
        return not self.errors


class FragmentParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.errors: list[str] = []
        self.stack: list[tuple[str, bool]] = []
        self.root_tags: list[str] = []

    def handle_decl(self, decl: str) -> None:
        self.errors.append(f"正文不能包含声明：<!{decl}>")

    def handle_comment(self, data: str) -> None:
        return

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self._handle_start(tag.lower(), attrs, closes_immediately=True)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self._handle_start(tag.lower(), attrs, closes_immediately=tag.lower() in VOID_TAGS)

    def _handle_start(
        self,
        tag: str,
        attrs: list[tuple[str, str | None]],
        *,
        closes_immediately: bool,
    ) -> None:
        if not self.stack:
            self.root_tags.append(tag)

        if tag in FORBIDDEN_TAGS:
            self.errors.append(f"禁止标签：<{tag}>")
        elif tag not in ALLOWED_TAGS:
            self.errors.append(f"未允许的标签：<{tag}>")

        attr_map: dict[str, str] = {}
        allowed_attrs = ATTRIBUTE_ALLOWLIST.get(tag, {"style"})
        for raw_name, raw_value in attrs:
            name = raw_name.lower()
            value = raw_value or ""
            if name.startswith("on"):
                self.errors.append(f"<{tag}> 不能包含事件属性 {name}")
            elif name not in allowed_attrs:
                self.errors.append(f"<{tag}> 不能包含属性 {name}")
            attr_map[name] = value

        if tag in STYLE_REQUIRED_TAGS and not attr_map.get("style", "").strip():
            self.errors.append(f"<{tag}> 缺少非空内联 style")

        style = attr_map.get("style", "")
        if style:
            self._validate_style(tag, style)

        if tag == "span" and "leaf" in attr_map and attr_map["leaf"] != "":
            self.errors.append('<span leaf> 的 leaf 值必须为空字符串')

        if tag == "a":
            href = attr_map.get("href", "")
            if not HTTP_LINK.match(href):
                self.errors.append("<a> 的 href 必须是完整的 http:// 或 https:// 地址")

        is_leaf = tag == "span" and "leaf" in attr_map
        if not closes_immediately:
            self.stack.append((tag, is_leaf))

    def _validate_style(self, tag: str, style: str) -> None:
        for pattern, label in FORBIDDEN_STYLE_PATTERNS:
            if pattern.search(style):
                self.errors.append(f"<{tag}> 的 style 包含禁止项：{label}")

        for declaration in style.split(";"):
            declaration = declaration.strip()
            if not declaration or ":" not in declaration:
                continue
            property_name, value = (part.strip() for part in declaration.split(":", 1))
            if COLOR_PROPERTY.match(property_name):
                normalized = re.sub(r"\s+", "", value).lower()
                if normalized not in ALLOWED_COLOR_VALUES:
                    self.errors.append(
                        f"<{tag}> 的 {property_name} 必须使用 inherit、"
                        "currentColor 或 transparent"
                    )
            elif BORDER_PROPERTY.match(property_name):
                tokens = value.split()
                if any(not BORDER_TOKEN.match(token) for token in tokens):
                    self.errors.append(
                        f"<{tag}> 的 {property_name} 不能包含具体颜色"
                    )
            elif property_name.lower() == "font-family":
                if value.strip().lower() != "inherit":
                    self.errors.append(f"<{tag}> 的 font-family 必须使用 inherit")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in VOID_TAGS:
            return
        if not self.stack:
            self.errors.append(f"多余的结束标签：</{tag}>")
            return
        open_tag, _ = self.stack[-1]
        if open_tag != tag:
            self.errors.append(f"标签未正确闭合：<{open_tag}> 后遇到 </{tag}>")
            return
        self.stack.pop()

    def handle_data(self, data: str) -> None:
        if not data.strip():
            return
        if not self.stack:
            self.errors.append("根 <section> 外存在可见文字")
            return
        if not any(is_leaf for _, is_leaf in self.stack):
            parent = self.stack[-1][0]
            snippet = data.strip()[:24]
            self.errors.append(f"<{parent}> 中的文字未放入 <span leaf=\"\">：{snippet}")

    def finish(self) -> None:
        if self.stack:
            self.errors.append(f"存在未闭合标签：<{self.stack[-1][0]}>")
        if self.root_tags != ["section"]:
            self.errors.append("正文必须且只能有一个根 <section>")


def validate_html(html: str) -> ValidationResult:
    errors: list[str] = []

    for pattern, message in FORBIDDEN_DOCUMENT_PATTERNS:
        if pattern.search(html):
            errors.append(message)

    parser = FragmentParser()
    try:
        parser.feed(html)
        parser.close()
        parser.finish()
    except Exception as exc:  # HTMLParser should not hide malformed input.
        errors.append(f"HTML 解析失败：{exc}")

    errors.extend(parser.errors)
    unique_errors = tuple(dict.fromkeys(errors))
    return ValidationResult(errors=unique_errors)


def main() -> int:
    parser = argparse.ArgumentParser(description="校验微信公众号正文 HTML")
    parser.add_argument("file", type=Path, help="公众号正文 HTML 文件")
    args = parser.parse_args()

    try:
        html = args.file.read_text(encoding="utf-8")
    except OSError as exc:
        parser.error(str(exc))

    result = validate_html(html)
    print(f"公众号正文校验：{args.file}")
    if result.ok:
        print("通过：正文结构、内联样式和无图片规则均符合要求。")
        return 0

    print(f"发现 {len(result.errors)} 个问题：")
    for error in result.errors:
        print(f"- {error}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
