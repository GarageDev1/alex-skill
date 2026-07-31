# 步骤 1：提取源文档

## 提取源文档内容

先运行提取脚本，将源文档转为可引用的结构化内容：

```
python scripts/extract_source.py report.pdf --output-dir workdir/
python scripts/extract_source.py report.docx --output-dir workdir/
python scripts/extract_source.py page.html --output-dir workdir/
```

- PDF/DOCX → 通过 PaddleOCR API 提取为 `source_content.md` + `_source_images/`
- HTML → 原样复制为 `source_content.html`

提取完成后，阅读 `source_content.md`（或 `.html`）建立论证骨架。后续步骤 3（创建 JSON）和步骤 3b（数据审计）均基于此文件工作。

## 建立论证骨架

从提取的内容中识别以下维度的信息：

- 公司业务、产品线、技术路径
- 经营数据、行业位置、竞争格局
- 风险和验证点

## 区分内容性质

先区分"业务研究"与"投资判断"。

对含投资建议的章节，进一步拆分：
- 哪些观点是产业事实/数据描述（适合 `both`）
- 哪些是投资动作/配置建议（仅 `internal`）

## 数据时效

- 历史数据必须保留来源日期和口径。
- 金融、市场和时事材料按需核验时效性。

## 禁止项

- 禁止自行填写`meta.date` 和 `share.badge`。日期必须由用户显式要求后，通过 `python scripts/set_date.py` 设置，JSON 中始终省略这两个字段。
