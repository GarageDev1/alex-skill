# 步骤 0：检查依赖

运行依赖检查脚本：

```
python scripts/check_deps.py
```

脚本检查 Python 版本、包管理器（uv/pip）、所需 Python 包（playwright、Pillow、pdf2image）、Chromium 浏览器和品牌资产文件。

- 返回码 `0`：全部确认存在，继续下一步。
- 返回码 `1`：有未确认项。根据输出提示尝试安装缺失依赖（优先 `uv pip install`，回退 `pip install`）；Chromium 使用 `playwright install chromium`。安装后重新运行脚本确认通过。
