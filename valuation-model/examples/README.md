# 通用示例

这里不保存公司专属生成脚本。历史版本需要时从 Git 查看。

## 生成模型

准备符合 `valuation_kit/schema/input.schema.json` 的真实输入文件，然后运行：

```bash
python examples/build_template.py \
  --input /path/to/input.json \
  --output /path/to/model.xlsx
```

Python 调用方式：

```python
from build_template import build_model

build_model("/path/to/input.json", "/path/to/model.xlsx")
```

`build_template.py` 不含示意数字、公司数字或备用数据。输入缺字段、来源、日期或币种口径时，生成会直接停止。

## 两个 Python 文件

- `build_template.py`：通用函数和命令行入口。
- `build_kit.py`：旧导入路径的兼容层。已有 `import build_kit as K` 可以继续使用，新代码应直接导入 `valuation_kit`。

端到端测试使用仓库外的真实输入。设置 `VALUATION_TEST_DATA_DIR`，目录内必须有：

```text
000660.KS_input.json
09880.HK_input.json
2330.TW_input.json
```

缺少环境变量或任一文件时，测试会明确失败。

无图形界面的测试环境可设置 `VALUATION_RECALC_DISABLE_EXCEL=1`，此时 `--backend auto` 会跳过 Excel，使用 LibreOffice。
