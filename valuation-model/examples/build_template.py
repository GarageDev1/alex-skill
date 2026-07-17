#!/usr/bin/env python3
"""从真实 input.json 生成估值工作簿。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from valuation_kit import build_workbook, load_input


def build_model(input_path, output_path):
    """读取并校验输入，然后生成工作簿。"""
    return build_workbook(load_input(input_path), output_path)


def main():
    parser = argparse.ArgumentParser(description="从真实 input.json 生成估值模型")
    parser.add_argument("--input", required=True, help="输入 JSON 文件")
    parser.add_argument("--output", required=True, help="输出 xlsx 文件")
    args = parser.parse_args()
    output = build_model(args.input, args.output)
    print(output.resolve())


if __name__ == "__main__":
    main()
