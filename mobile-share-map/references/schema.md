# 手机转发长图 JSON 结构

使用 `output_policy` 明确选择单版或双版。中性信息整理使用 `single`；针对具体可交易标的且包含投资建议、评级、估值、目标价、仓位或交易判断时使用 `dual`。无法判断时先询问用户。

## Single 单版示例

```json
{
  "output_policy": "single",
  "title": "HBM扩产竞赛：谁在建厂、谁在抢单",
  "subtitle": "产能结构与扩产进度一图看清——这轮建厂潮决定明年内存条和手机存储贵不贵",
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
  "title": "XX公司AI芯片突围：自研架构能否撬动英伟达的定价权",
  "subtitle": "技术路线与客户进展全景梳理——云计算降本的下一张牌，影响你用的每一个AI产品的价格",
  "external_title": "XX公司AI芯片布局：自研架构与客户生态全景",
  "external_subtitle": "技术路线与合作进展一图看清——AI芯片的竞争格局如何影响云服务和智能产品的成本",
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

## 标题与副标题

标题和副标题必须同时服务两类读者：外行快速获得阅读动机，内行通过关键词判断是否需要阅读。

### 标题（title）

回答"是什么"。保留核心行业术语作为内行的筛选信号，但用"谁/什么/怎样"的动作句式替代名词堆叠，让外行也能读进去。

- 好：`HBM扩产竞赛：谁在建厂、谁在抢单`（内行看到"HBM扩产"秒判相关性，外行看到"谁在建厂"有画面感）
- 坏：`HBM主要厂商扩产图谱`（名词堆叠，外行没有阅读动机）

### 副标题（subtitle）

回答"和我有什么关系"+"看了有什么用"。结构为：**前半句点明覆盖维度**（内行确认深度）+ **后半句用传导链落到日常生活**（外行获得动机）。

**传导链思维**：从行业话题出发，沿因果链条向下游推导，直到触达普通人可感知的生活场景——消费品价格、日常使用的产品或服务、就业机会等。

传导链示例：
- HBM扩产 → 存储芯片供需 → 内存/硬盘价格 → **明年你换手机和加内存大概率要多掏钱**
- AI芯片竞争 → 云计算成本 → AI产品定价 → **你的AI会员快要涨价了**
- AI算力融资 → 云计算成本 → AI产品定价 → **你的ChatGPT会员可能要涨价了**
- 光伏产能过剩 → 组件价格战 → **今年装家用光伏能省下一大笔钱**
- 航运运力紧张 → 运费上涨 → **你的海淘包裹运费已经在涨了**

副标题不需要走完整条链，只需比源文档的行业视角再推一步，用一句大白话陈述落到读者能感知的生活场景。

- 好：`产能结构与扩产进度一图看清，明年你换手机和加内存大概率要多掏钱`
- 坏：`产能结构、项目进度与信息口径`（纯维度罗列，没有回答"和我有什么关系"）

### 禁止项

- **指令回声**：禁止"不需要X，只需要Y""不是X，而是Y""X的本质是Y"等教科书句式，这是 Agent 把内部推理过程直接输出为标题的典型痕迹。
- **设问**：标题和副标题禁止出现"为什么""如何""怎样""是否""吗""贵不贵""划不划算"等疑问词和正反问句式。
- **破折号**：禁止使用"——"连接前后半句，用逗号。
- **副标题后半句必须是大白话陈述句**：一个普通人会对朋友说的话，不是论文摘要。
  - 好："你的ChatGPT会员可能要涨价了"
  - 坏："AI投资的减速比崩塌更值得每个人关注"（像社论，不像人话）
  - 坏："算力过剩的账单由AI用户和科技股持有者买单"（强行塞两个对象，像新闻标题模板）

### Dual 模式

`external_title` 和 `external_subtitle` 遵循相同的生成规则，但不得包含投资敏感内容。传导链的终点应落在产品、技术或消费场景，而非股价或投资收益。

## 字段约束

- `meta.date` 和 `share.badge` 禁止 Agent 自行填写，在 JSON 中始终省略。日期必须由用户显式要求后通过 `python scripts/set_date.py <input.json> [--date YYYY-MM-DD]` 设置。渲染器启动时会检查日期是否已设置，未设置则拒绝渲染。
- `output_policy` 必须为 `single` 或 `dual`。省略时默认的 `--mode auto` 会停止渲染，要求先完成策略判断。
- `dual` 模式的 `audience` 只能为 `both` 或 `internal`。省略时按 `internal` 处理，防止误外发。
- `tag` 是左侧浅色标签的短章节名，建议 6-12 个汉字；`tone` 可使用 `blue`、`amber`、`green`、`pink`、`cyan`、`purple`、`orange` 或 `gray`，省略时按顺序循环。
- `blocks` 是章节内的视觉块，目前支持 `type: "table"`。表格使用 `title`、`headers` 和 `rows`；表头数量须与每行单元格数一致，限 2-4 列，每个单元格仅放简短事实或判断。
- `dual` 模式的每个 `blocks` 条目必须显式写 `audience`。External 仅保留 `audience: "both"` 的表格，并会逐格检查敏感词；任何带评级、估值、价格、交易建议或预测性财务信息的表一律标为 `internal`。
- `external_title` 和 `external_subtitle` 应为业务研究表述；不提供时才退回到内部标题/副标题，并接受敏感词检查。
- `style.safe_top_px` 限制在 `72-240`；`style.accent` 仅接受 `#RGB` 或 `#RRGGBB`。
- 品牌背景、文案和二维码默认读取 `assets/brand.json`。`footer_intro` 支持用 JSON 转义 `\n` 指定必须保留的真实换行，默认机构介绍固定为两行。 当 `share.override_brand` 不为 `true` 时，`brand.json` 中已存在的字段优先，即使其值为空字符串或空数组；因此默认的 `footer_points: []` 会压制输入 JSON 中遗留的“人 / 事 / 钱”分点。
- 只有设置 `share.override_brand: true`，输入 `share` 中显式出现的品牌字段才覆盖 `brand.json`；未写入的字段仍继承品牌配置。该规则适用于 `masthead_text`、`footer_title`、`footer_intro`、`footer_points`、`footer_background`、`qr_image` 与 `qr_label`，并允许用空字符串或空数组明确清空字段。
- `share.footer_background` 是不含文字和二维码的 1080×400px 底栏背景图片路径；默认使用 `assets/footer-tech-blue.png`。底栏文字与二维码始终由渲染器独立分层，正文与底栏之间固定保留 70px 白色间隔，底栏与间隔合计占用 470px。
- `share.qr_image` 最终解析为空或文件不存在时生成二维码预留位；默认品牌配置使用 `assets/zhifujie-qr.png`，顶部与底部共用同一图片。
- External 不使用卡片 `source`，避免在来源文本中意外带出敏感投资术语。
- `single` 不渲染任何受众标签；`dual` 的内部图由渲染器统一显示“内部研究版”，External不显示受众标签。
- External 忽略 `share.badge`，右上角固定显示 `meta.date`。

## 研报预览横条

必须生成缩略图并在 JSON 中加入 `docx_preview` 字段，在正文与底栏之间插入研报页面缩略图预览横条：

```json
{
  "docx_preview": {
    "dir": "_page_thumbs",
    "pages": 4,
    "thumb_width": 240,
    "total_pages": 42
  }
}
```

- `dir`：页面缩略图目录路径，相对于输入 JSON 所在目录。目录中需包含 `page_01.png`、`page_02.png` 等按页码命名的 PNG 文件。
- `pages`：展示的页面数量，默认 `4`。渲染器每行最多显示 4 张缩略图，超出自动换行；例如 `pages: 8` 会显示为两行。
- `thumb_width`：缩略图嵌入时缩放到的宽度（像素），默认 `240`。仅影响嵌入尺寸，不改变显示尺寸。
- `total_pages`：可选，研报总页数。省略时自动从同目录下的 PDF 文件读取；无 PDF 时退回到目录内的 PNG 文件数。

预览横条在 Internal 和 External 两版中均显示。缩略图缩到 210px 宽后页面文字不可读，不构成合规风险。

渲染器会在启动时检查缩略图是否就绪（`docx_preview` 字段和实际文件），缺少时拒绝渲染。
