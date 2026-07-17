from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = ROOT / "examples"


class RealInputEndToEndTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        directory = os.environ.get("VALUATION_TEST_DATA_DIR")
        if not directory:
            raise AssertionError("VALUATION_TEST_DATA_DIR 未设置")
        cls.data_dir = Path(directory)
        cls.inputs = [cls.data_dir / f"{ticker}_input.json" for ticker in ("000660.KS", "09880.HK", "2330.TW")]
        missing = [str(path) for path in cls.inputs if not path.is_file()]
        if missing:
            raise AssertionError(f"缺少真实测试输入：{missing}")

        spec = importlib.util.spec_from_file_location("build_template", EXAMPLES / "build_template.py")
        cls.builder = importlib.util.module_from_spec(spec)
        assert spec and spec.loader
        spec.loader.exec_module(cls.builder)

    def test_three_real_inputs_recalculate_and_pass_validator(self):
        with tempfile.TemporaryDirectory() as tmp:
            for input_path in self.inputs:
                output = Path(tmp) / f"{input_path.stem}.xlsx"
                self.builder.build_model(input_path, output)
                env = {**os.environ, "VALUATION_RECALC_DISABLE_EXCEL": "1"}
                subprocess.run([sys.executable, str(ROOT / "scripts" / "recalc.py"), "--backend", "auto", str(output)], check=True, capture_output=True, text=True, env=env)
                result = subprocess.run([sys.executable, str(ROOT / "scripts" / "validate_valuation.py"), str(output)], check=True, capture_output=True, text=True)
                payload = json.loads(result.stdout)
                self.assertEqual(payload["verdict"], "PASS", payload)
                self.assertEqual(payload["summary"], {"pass": 13, "warn": 0, "fail": 0})


if __name__ == "__main__":
    unittest.main()
