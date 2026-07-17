from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import valuation_kit as kit


class ValuationKitTests(unittest.TestCase):
    def test_currency_and_per_share_conversions(self):
        self.assertAlmostEqual(kit.local_billions_to_usd_billions(72, 7.2), 10)
        self.assertAlmostEqual(kit.usd_per_share(10, 1000), 10)
        self.assertAlmostEqual(kit.market_cap_usd_billions(78, 1000, 7.8), 10)

    def test_multiple_methods(self):
        self.assertAlmostEqual(kit.pe_price(10, 1000, 20, 7.8), 1560)
        self.assertAlmostEqual(kit.pb_price(20, 1000, 4, 7.8), 624)
        self.assertAlmostEqual(kit.ev_sales_price(10, 2, 1000, 5, 7.8), 374.4)
        self.assertAlmostEqual(kit.dcf_value([10, 10], 0.10, 0.02), 122.7272727273)
        self.assertEqual(kit.sotp_value([{"value": 10}, {"metric": 5, "multiple": 3}], 2), 23)

    def test_scenario_switch(self):
        self.assertEqual(kit.select_case({"Bear": 1, "Base": 2, "Bull": 3}, "Base"), 2)
        self.assertEqual(kit.nested_if_switch("B2"), '=IF(B2="Bear",1,IF(B2="Base",2,3))')

    def test_compatibility_module_reexports_public_api(self):
        path = ROOT / "examples" / "build_kit.py"
        spec = importlib.util.spec_from_file_location("legacy_build_kit", path)
        module = importlib.util.module_from_spec(spec)
        assert spec and spec.loader
        spec.loader.exec_module(module)
        for name in ("load_input", "validate_input", "build_workbook", "write_cover", "pe_price"):
            self.assertIs(getattr(module, name), getattr(kit, name))


if __name__ == "__main__":
    unittest.main()
