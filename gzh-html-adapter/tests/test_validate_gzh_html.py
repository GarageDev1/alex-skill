from __future__ import annotations

import sys
import unittest
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from validate_gzh_html import validate_html  # noqa: E402


ROOT_STYLE = (
    "max-width:100%;box-sizing:border-box;background:#f7f8fb;"
    "color:#263238;font-family:inherit;"
)
BLOCK_STYLE = (
    "margin:1em 0;color:#263238;background:#ffffff;"
    "border:1px solid #d7e0e8;box-sizing:border-box;"
)
TEXT_STYLE = "margin:0;line-height:1.75;color:#263238;"
TABLE_STYLE = (
    "width:100%;border-collapse:collapse;table-layout:fixed;color:#263238;"
)
CELL_STYLE = (
    "padding:.5em;border:1px solid #c7d4e3;vertical-align:top;"
    "color:#263238;background:#ffffff;"
)
ROW_STYLE = "color:#263238;"


def leaf(text: str) -> str:
    return f'<span leaf="">{text}</span>'


def table(rows: list[tuple[str, str]]) -> str:
    body = "".join(
        f'<tr style="{ROW_STYLE}">'
        f'<td style="{CELL_STYLE}">{leaf(label)}</td>'
        f'<td style="{CELL_STYLE}">{leaf(value)}</td>'
        "</tr>"
        for label, value in rows
    )
    return (
        f'<table style="{TABLE_STYLE}">'
        f'<tbody style="{ROW_STYLE}">{body}</tbody>'
        "</table>"
    )


def compliant_fragment() -> str:
    kpi = (
        f'<section style="{BLOCK_STYLE}">'
        f'<p style="{TEXT_STYLE}">{leaf("关键指标")}</p>'
        f'{table([("指标", "数值")])}'
        "</section>"
    )
    data_table = (
        f'<section style="{BLOCK_STYLE}">'
        f'<p style="{TEXT_STYLE}">{leaf("数据表")}</p>'
        f'{table([("类别", "数值")])}'
        "</section>"
    )
    horizontal_bar = (
        f'<section style="{BLOCK_STYLE}">'
        f'<p style="{TEXT_STYLE}">{leaf("横向条形图")}</p>'
        f'<section style="width:100%;border:1px solid #b7c7d9;'
        f'background:#eef3f8;box-sizing:border-box;color:#263238;">'
        f'<section style="width:50%;height:1em;background:#2f6f8f;'
        f'color:#ffffff;">{leaf(" ")}</section>'
        "</section>"
        f'<p style="{TEXT_STYLE}">{leaf("类别：数值")}</p>'
        "</section>"
    )
    vertical_bar = (
        f'<section style="{BLOCK_STYLE}">'
        f'<p style="{TEXT_STYLE}">{leaf("纵向条形图")}</p>'
        f'<table style="{TABLE_STYLE}"><tbody style="{ROW_STYLE}">'
        f'<tr style="{ROW_STYLE}"><td style="{CELL_STYLE}">'
        f'<section style="height:2em;background:#2f6f8f;'
        f'color:#ffffff;">{leaf(" ")}</section>'
        f'<p style="{TEXT_STYLE}">{leaf("时期：数值")}</p>'
        "</td></tr></tbody></table></section>"
    )
    stacked_bar = (
        f'<section style="{BLOCK_STYLE}">'
        f'<p style="{TEXT_STYLE}">{leaf("堆叠条形图")}</p>'
        f'<section style="width:100%;font-size:0;color:#263238;">'
        f'<span style="display:inline-block;width:40%;height:1em;'
        f'background:#2f6f8f;"></span>'
        f'<span style="display:inline-block;width:60%;height:1em;'
        f'background:#74a6b8;"></span>'
        "</section>"
        f'<p style="{TEXT_STYLE}">{leaf("分项及数值")}</p>'
        "</section>"
    )
    progress = (
        f'<section style="{BLOCK_STYLE}">'
        f'<p style="{TEXT_STYLE}">{leaf("进度图：实际值与目标值")}</p>'
        f'<section style="width:100%;border:1px solid #b7c7d9;'
        f'background:#eef3f8;box-sizing:border-box;color:#263238;">'
        f'<section style="width:75%;height:1em;background:#477a5b;'
        f'color:#ffffff;">{leaf(" ")}</section>'
        "</section></section>"
    )
    timeline = (
        f'<section style="{BLOCK_STYLE}">'
        f'<p style="{TEXT_STYLE}">{leaf("时间线")}</p>'
        f'{table([("日期", "事件")])}'
        "</section>"
    )
    link = (
        f'<p style="{TEXT_STYLE}">'
        f'<a href="https://example.com/source" '
        f'style="color:inherit;text-decoration:underline;">'
        f'{leaf("来源")}</a></p>'
    )
    return (
        f'<section style="{ROOT_STYLE}">'
        f"{kpi}{data_table}{horizontal_bar}{vertical_bar}"
        f"{stacked_bar}{progress}{timeline}{link}"
        "</section>"
    )


class ValidateWechatHtmlTests(unittest.TestCase):
    def test_accepts_all_supported_native_chart_structures(self) -> None:
        result = validate_html(compliant_fragment())
        self.assertTrue(result.ok, result.errors)

    def test_rejects_all_image_and_active_content_paths(self) -> None:
        forbidden = {
            "img": '<img src="chart.png">',
            "picture": "<picture></picture>",
            "svg": "<svg></svg>",
            "canvas": "<canvas></canvas>",
            "script": "<script></script>",
            "style": "<style></style>",
            "external-style": '<link rel="stylesheet" href="theme.css">',
            "button": "<button>复制</button>",
            "css-url": '<section style="background:url(chart.png);"></section>',
            "background-image": (
                '<section style="background-image:linear-gradient(#fff,#eee);">'
                "</section>"
            ),
            "data-uri": (
                '<section style="background-image:url(data:image/png;base64,AA);">'
                "</section>"
            ),
            "forbidden-position": (
                '<section style="position:absolute;color:#222222;"></section>'
            ),
        }
        for name, payload in forbidden.items():
            with self.subTest(name=name):
                html = (
                    f'<section style="{ROOT_STYLE}">'
                    f'<section style="{BLOCK_STYLE}">{payload}</section>'
                    "</section>"
                )
                self.assertFalse(validate_html(html).ok)

    def test_rejects_missing_inline_style(self) -> None:
        html = f'<section style="{ROOT_STYLE}"><p>{leaf("正文")}</p></section>'
        self.assertFalse(validate_html(html).ok)

    def test_rejects_unwrapped_text(self) -> None:
        html = f'<section style="{ROOT_STYLE}"><p style="{TEXT_STYLE}">正文</p></section>'
        self.assertFalse(validate_html(html).ok)

    def test_accepts_contextual_colors_shadows_and_internal_dark_block(self) -> None:
        html = (
            '<section style="max-width:100%;box-sizing:border-box;'
            'background:linear-gradient(135deg,#ffffff 0%,#eef4ff 100%);'
            'font-family:inherit;color:#263238;">'
            '<section style="margin:1em 0;padding:1em;background:#17202a;'
            'color:#ffffff;border:1px solid #314152;'
            'box-shadow:0 2px 8px rgba(0,0,0,.15);">'
            f'{leaf("局部深色强调")}</section></section>'
        )
        result = validate_html(html)
        self.assertTrue(result.ok, result.errors)

    def test_accepts_light_solid_and_gradient_root_backgrounds(self) -> None:
        backgrounds = (
            "#f8f6f1",
            "rgb(246,248,252)",
            "hsl(210,40%,97%)",
            "linear-gradient(135deg,#ffffff 0%,rgb(238,244,255) 100%)",
            "radial-gradient(circle at top,hsl(210,100%,98%),#f5f0ff)",
            (
                "repeating-linear-gradient(45deg,#ffffff 0,#ffffff 8px,"
                "#f2f5f8 8px,#f2f5f8 16px)"
            ),
        )
        for background in backgrounds:
            with self.subTest(background=background):
                html = (
                    '<section style="max-width:100%;box-sizing:border-box;'
                    f'background:{background};font-family:inherit;"></section>'
                )
                result = validate_html(html)
                self.assertTrue(result.ok, result.errors)

    def test_accepts_light_background_color_declaration(self) -> None:
        html = (
            '<section style="max-width:100%;background-color:#f4f7f9;'
            'font-family:inherit;"></section>'
        )
        result = validate_html(html)
        self.assertTrue(result.ok, result.errors)

    def test_rejects_missing_dark_and_unparseable_root_backgrounds(self) -> None:
        styles = (
            "max-width:100%;box-sizing:border-box;font-family:inherit",
            "max-width:100%;background:#111827;font-family:inherit",
            (
                "max-width:100%;background:linear-gradient(90deg,#ffffff,#334155);"
                "font-family:inherit"
            ),
            (
                "max-width:100%;background:linear-gradient(90deg,#111827,#334155);"
                "background-color:#ffffff;font-family:inherit"
            ),
            (
                "max-width:100%;background:color-mix(in srgb,white 80%,blue);"
                "font-family:inherit"
            ),
        )
        for style in styles:
            with self.subTest(style=style):
                self.assertFalse(validate_html(f'<section style="{style}"></section>').ok)

    def test_rejects_non_inherited_font_family(self) -> None:
        html = (
            f'<section style="{ROOT_STYLE}">'
            f'<p style="{TEXT_STYLE}font-family:Arial;">{leaf("正文")}</p>'
            "</section>"
        )
        self.assertFalse(validate_html(html).ok)

    def test_rejects_document_shell_and_multiple_roots(self) -> None:
        document = (
            "<!DOCTYPE html><html><body>"
            f'<section style="{ROOT_STYLE}"></section>'
            "</body></html>"
        )
        self.assertFalse(validate_html(document).ok)

        multiple = (
            f'<section style="{ROOT_STYLE}"></section>'
            f'<section style="{ROOT_STYLE}"></section>'
        )
        self.assertFalse(validate_html(multiple).ok)

    def test_rejects_relative_or_active_links(self) -> None:
        for href in ("/relative", "#anchor", "javascript:alert(1)", "data:text/plain,x"):
            with self.subTest(href=href):
                html = (
                    f'<section style="{ROOT_STYLE}">'
                    f'<a href="{href}" style="color:inherit;">{leaf("链接")}</a>'
                    "</section>"
                )
                self.assertFalse(validate_html(html).ok)


if __name__ == "__main__":
    unittest.main()
