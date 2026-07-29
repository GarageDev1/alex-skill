# 第一步：环境检查

运行环境依赖检测脚本：

```bash
python "<skill-dir>/scripts/preflight.py"
```

脚本会检测 Node.js、Playwright 和浏览器。

- `[FOUND]` 表示确认存在。
- `[UNCERTAIN]` 表示无法确认——需要自行判断是否可以继续。

## 前提

- 需要 Node.js 和 Playwright 及 Chromium 浏览器。渲染器会在当前工作区和 SKILL 的上级工作区搜索 `playwright`。
- 如果关键依赖缺失且无法解决，终止流程并告知用户。
