# 步骤 5：渲染

运行渲染脚本：

```
python scripts/render_mobile_share.py input.json --output-dir outputs
```

## 模式说明

默认 `--mode auto`，根据 JSON 中的 `output_policy` 自动选择：
- `single` → 输出 `input.png`
- `dual` → 输出 `input-internal.png` 与 `input-external.png`，并保留对应 HTML 以便审阅

JSON 缺少 `output_policy` 时渲染器会停止。应返回步骤 2 重新判定，必要时询问用户。

## 调试与覆盖

- `--mode single|both|internal|external`：仅用于明确覆盖或调试，不要用来绕过输出策略判断。
- `--html-only`：仅生成 HTML，跳过 PNG。仅用于调试，正式交付不得使用。
