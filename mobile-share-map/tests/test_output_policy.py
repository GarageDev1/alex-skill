import importlib.util
import json
import struct
import tempfile
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "render_mobile_share.py"
SPEC = importlib.util.spec_from_file_location("render_mobile_share", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


def neutral_data():
    return {
        "output_policy": "single",
        "title": "HBM扩产图谱",
        "subtitle": "base die、IPO进度与产能预测",
        "meta": {"source": "公开资料", "date": "2026-07-13"},
        "sections": [
            {
                "title": "供给判断",
                "tag": "供给判断",
                "cards": [
                    {
                        "label": "产能口径",
                        "text": "只看DRAM投片会高估短期供给，这是运营口径判断。",
                    }
                ],
            }
        ],
    }


class OutputPolicyTests(unittest.TestCase):
    def test_brand_defaults_match_research_report_cta(self):
        brand_path = SCRIPT.parents[1] / "assets" / "brand.json"
        brand = json.loads(brand_path.read_text(encoding="utf-8"))
        self.assertEqual(brand["footer_title"], "完整研报加入智富界交流群")
        self.assertEqual(brand["qr_label"], "扫码获取完整研报")
        self.assertEqual(brand["footer_points"], [])
        self.assertTrue((SCRIPT.parents[1] / brand["footer_background"]).is_file())

    def test_footer_uses_layered_background_and_fixed_height(self):
        html = MODULE.render_html(MODULE.single_variant(neutral_data()), "single")
        self.assertIn('class="footer-background"', html)
        self.assertIn("height:536px", html)
        self.assertIn("完整研报加入智富界交流群", html)
        self.assertIn('<div class="header-qr-label">扫码获取完整研报</div>', html)
        self.assertIn('<div class="qr-label">扫码获取完整研报</div>', html)
        self.assertRegex(html, r'<img class="header-qr-image"[^>]+alt="扫码获取完整研报">')
        self.assertRegex(html, r'<img class="qr-image"[^>]+alt="扫码获取完整研报">')
        self.assertNotIn('class="bottom-safe-area"', html)
        self.assertNotIn("AI IP 与 AI 员工", html)

    def test_footer_points_override_remains_compatible(self):
        data = neutral_data()
        data["share"] = {
            "override_brand": True,
            "footer_points": [{"label": "人", "text": "兼容说明"}],
        }
        html = MODULE.render_html(MODULE.single_variant(data), "single")
        self.assertIn("兼容说明", html)

    def test_brand_empty_points_suppress_unoverridden_input_points(self):
        data = neutral_data()
        data["share"] = {"footer_points": [{"label": "旧", "text": "不应出现"}]}
        html = MODULE.render_html(MODULE.single_variant(data), "single")
        self.assertNotIn("不应出现", html)

    def test_render_png_crops_to_page_height(self):
        html = MODULE.render_html(MODULE.single_variant(neutral_data()), "single")
        with tempfile.TemporaryDirectory() as temp_dir:
            html_path = Path(temp_dir) / "short.html"
            png_path = Path(temp_dir) / "short.png"
            html_path.write_text(html, encoding="utf-8")
            MODULE.render_png(html_path, png_path)
            png = png_path.read_bytes()
            width, height = struct.unpack(">II", png[16:24])
        self.assertEqual(width, 1080)
        self.assertLess(height, 1920)
        self.assertGreaterEqual(height, 536)

    def test_single_allows_neutral_industry_language(self):
        variant = MODULE.single_variant(neutral_data())
        html = MODULE.render_html(variant, "single")
        self.assertNotIn("内部研究版", html)

    def test_single_rejects_explicit_investment_advice(self):
        data = neutral_data()
        data["sections"][0]["cards"][0]["text"] = "评级买入，目标价100元。"
        with self.assertRaisesRegex(ValueError, "output_policy: dual"):
            MODULE.single_variant(data)

    def test_dual_labels_internal_and_filters_external(self):
        data = {
            "output_policy": "dual",
            "title": "标的研究",
            "external_title": "公司业务图谱",
            "meta": {"date": "2026-07-13"},
            "sections": [
                {
                    "title": "业务与判断",
                    "tag": "业务判断",
                    "audience": "both",
                    "cards": [
                        {"label": "业务事实", "text": "产品已经量产。", "audience": "both"},
                        {"label": "内部评级", "text": "买入，目标价100元。", "audience": "internal"},
                    ],
                }
            ],
        }
        internal_html = MODULE.render_html(MODULE.internal_variant(data), "internal")
        external_html = MODULE.render_html(MODULE.external_variant(data), "external")
        self.assertIn("内部研究版", internal_html)
        self.assertNotIn("内部研究版", external_html)
        self.assertNotIn("目标价100元", external_html)


if __name__ == "__main__":
    unittest.main()
