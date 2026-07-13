import importlib.util
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
