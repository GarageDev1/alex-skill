from __future__ import annotations

from copy import deepcopy
import json
import os
from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from valuation_kit import InputValidationError, load_input, validate_input


class InputValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        directory = os.environ.get("VALUATION_TEST_DATA_DIR")
        if not directory:
            raise AssertionError("VALUATION_TEST_DATA_DIR 未设置")
        cls.fixture = Path(directory) / "2330.TW_input.json"
        if not cls.fixture.is_file():
            raise AssertionError(f"缺少真实测试输入：{cls.fixture}")
        cls.data = json.loads(cls.fixture.read_text(encoding="utf-8"))

    def test_valid_input(self):
        self.assertEqual(load_input(self.fixture)["company"]["ticker"], "2330.TW")

    def test_missing_source_is_rejected(self):
        data = deepcopy(self.data)
        del data["sources"][data["history"]["revenue"]["source_id"]]
        with self.assertRaisesRegex(InputValidationError, "sources 缺少引用项"):
            validate_input(data)

    def test_misaligned_years_are_rejected(self):
        data = deepcopy(self.data)
        data["years"]["forecast"][0] = data["years"]["historical"][-1]
        with self.assertRaisesRegex(InputValidationError, "不能重叠或错位"):
            validate_input(data)

    def test_currency_mismatch_is_rejected(self):
        data = deepcopy(self.data)
        data["company"]["price"]["currency"] = "USD"
        with self.assertRaisesRegex(InputValidationError, "股价币种"):
            validate_input(data)

    def test_one_method_is_rejected(self):
        data = deepcopy(self.data)
        data["valuation_methods"] = data["valuation_methods"][:1]
        with self.assertRaises(InputValidationError):
            validate_input(data)

    def test_wrong_field_type_is_rejected(self):
        data = deepcopy(self.data)
        data["history"]["revenue"]["values"][0] = "56.81"
        with self.assertRaises(InputValidationError):
            validate_input(data)


if __name__ == "__main__":
    unittest.main()
