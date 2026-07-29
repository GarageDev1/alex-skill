# 步骤 4：生成研报页面缩略图

缩略图的作用是吸引阅读的用户加群。除非用户显式要求省略，否则必须生成。

## 按源格式处理

- **PDF**：直接用 `pdf2image` 转图。
- **DOCX**：先通过 `win32com` 导出 PDF 再用 `pdf2image` 转图，或使用 LibreOffice headless。
- **其他格式**：自行判断什么样的缩略图合适且能吸引用户，把推荐选项告诉用户并征求意见。

## 输出规范

将生成的缩略图放入输入 JSON 同目录下的 `_page_thumbs/` 子目录，文件命名为 `page_01.png`、`page_02.png` 等。

在 JSON 中加入 `docx_preview` 字段：

```json
{
  "docx_preview": {
    "dir": "_page_thumbs",
    "pages": 4
  }
}
```

`pages` 为实际生成的缩略图数量。渲染器每行最多显示 4 张，超出自动换行。

详见 `references/schema.md` 的"研报预览横条"章节。
