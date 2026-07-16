# 手机转发长图 JSON 结构

使用 `output_policy` 明确选择单版或双版。中性信息整理使用 `single`；针对具体可交易标的且包含投资建议、评级、估值、目标价、仓位或交易判断时使用 `dual`。无法判断时先询问用户。

## Single 单版示例

```json
{
  "output_policy": "single",
  "title": "HBM主要厂商扩产图谱",
  "subtitle": "产能结构、项目进度与信息口径",
  "meta": {"source": "公开资料整理", "date": "YYYY-MM-DD"},
  "sections": [
    {
      "title": "行业产能格局",
      "tag": "产能格局",
      "tone": "blue",
      "cards": [
        {"label": "供给结构", "text": "以统一口径呈现公开产能信息。"}
      ]
    }
  ]
}
```

`single` 不要求 `audience`，输出 `input.png`。如果输入含评级、估值、目标价、交易建议等投资敏感内容，渲染器会拒绝生成，提示改用 `dual` 或重新确认。

## Dual 双版示例

External 仅渲染显式标为 `audience: "both"` 的内容，并再通过敏感关键词过滤。

```json
{
  "output_policy": "dual",
  "title": "标的研究标题",
  "subtitle": "内部内容可包含完整投资判断，但标题不手写版本标签。",
  "external_title": "业务研究标题",
  "external_subtitle": "仅概述公司业务、产品、技术与公开历史事实。",
  "meta": {
    "source": "公开披露与公司资料",
    "date": "YYYY-MM-DD",
    "analyst": "可选"
  },
  "style": {
    "safe_top_px": 132,
    "accent": "#e45b3f"
  },
  "share": {
    "override_brand": true,
    "masthead_text": "智富界 · AI炒股达人都在这",
    "badge": "YYYY.MM.DD",
    "footer_title": "完整研报加入智富界交流群",
    "footer_intro": "智富界是一个聚焦AI产业、创业与投资的研究平台，帮助企业及用户看懂AI、用好AI、投资AI。",
    "footer_points": [],
    "footer_background": "assets/footer-tech-blue.png",
    "qr_image": "assets/zhifujie-qr.png",
    "qr_label": "扫码获取完整研报"
  },
  "sections": [
    {
      "title": "公司业务与产品布局",
      "tag": "产品与业务",
      "tone": "blue",
      "audience": "both",
      "cards": [
        {
          "label": "产品结构",
          "text": "公司覆盖多个产品线，形成从核心能力到应用场景的协同。",
          "audience": "both",
          "source": "年报，YYYY-MM-DD"
        },
        {
          "label": "内部估值判断",
          "text": "仅供内部讨论的估值、评级或目标价格判断。",
          "audience": "internal"
        }
      ],
      "blocks": [
        {
          "type": "table",
          "title": "关键数据对照",
          "headers": ["维度", "数据", "说明"],
          "rows": [
            ["业务结构", "公开口径", "用于解释产品与经营特征"]
          ],
          "audience": "both"
        }
      ]
    }
  ]
}
```

## 字段约束

- `output_policy` 必须为 `single` 或 `dual`。省略时默认的 `--mode auto` 会停止渲染，要求先完成策略判断。
- `dual` 模式的 `audience` 只能为 `both` 或 `internal`。省略时按 `internal` 处理，防止误外发。
- `tag` 是左侧浅色标签的短章节名，建议 6-12 个汉字；`tone` 可使用 `blue`、`amber`、`green`、`pink`、`cyan`、`purple`、`orange` 或 `gray`，省略时按顺序循环。
- `blocks` 是章节内的视觉块，目前支持 `type: "table"`。表格使用 `title`、`headers` 和 `rows`；表头数量须与每行单元格数一致，限 2-4 列，每个单元格仅放简短事实或判断。
- `dual` 模式的每个 `blocks` 条目必须显式写 `audience`。External 仅保留 `audience: "both"` 的表格，并会逐格检查敏感词；任何带评级、估值、价格、交易建议或预测性财务信息的表一律标为 `internal`。
- `external_title` 和 `external_subtitle` 应为业务研究表述；不提供时才退回到内部标题/副标题，并接受敏感词检查。
- `style.safe_top_px` 限制在 `72-240`；`style.accent` 仅接受 `#RGB` 或 `#RRGGBB`。
- 品牌背景、文案和二维码默认读取 `assets/brand.json`。当 `share.override_brand` 不为 `true` 时，`brand.json` 中已存在的字段优先，即使其值为空字符串或空数组；因此默认的 `footer_points: []` 会压制输入 JSON 中遗留的“人 / 事 / 钱”分点。
- 只有设置 `share.override_brand: true`，输入 `share` 中显式出现的品牌字段才覆盖 `brand.json`；未写入的字段仍继承品牌配置。该规则适用于 `masthead_text`、`footer_title`、`footer_intro`、`footer_points`、`footer_background`、`qr_image` 与 `qr_label`，并允许用空字符串或空数组明确清空字段。
- `share.footer_background` 是不含文字和二维码的底栏背景图片路径；默认使用 `assets/footer-tech-blue.png`。底栏文字与二维码始终由渲染器独立分层。
- `share.qr_image` 最终解析为空或文件不存在时生成二维码预留位；默认品牌配置使用 `assets/zhifujie-qr.png`，顶部与底部共用同一图片。
- External 不使用卡片 `source`，避免在来源文本中意外带出敏感投资术语。
- `single` 不渲染任何受众标签；`dual` 的内部图由渲染器统一显示“内部研究版”，External不显示受众标签。
- External 忽略 `share.badge`，右上角固定显示 `meta.date`。
