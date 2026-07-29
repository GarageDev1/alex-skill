# 第六步：渲染

```bash
node "<skill-dir>/scripts/render.mjs" <input.md> --output <output-dir>
```

该命令输出编号的 1080×1440 PNG 文件和 `manifest.json`。会根据内容自动生成所需数量的正文页，并默认添加封面（设置 `cover: false` 可跳过）。使用 `--theme classic|finance|editorial|tech` 可预览或覆盖 front matter 中的主题而无需编辑输入文件。

渲染器返回非零退出码时视为真实失败。修复输入或运行环境；不得谎称图片已生成。
