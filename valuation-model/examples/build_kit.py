"""旧导入路径兼容层。新代码直接导入 valuation_kit。"""

from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from valuation_kit import *
from valuation_kit import __all__
