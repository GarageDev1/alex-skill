---
name: text-to-mindmap
description: 把文本内容制作成竖向逻辑树思维导图(PNG)。当用户说"把这段内容做成思维导图""文本转思维导图""生成一张思维导图""帮我画个思维导图""整理成思维导图图片""做成 mindmap/逻辑树图",或给一段资料/大纲/文章要求可视化成思维导图时使用。结构化大纲可直接转换;散文或长文先整理为大纲;图片素材先读取为文本。输出一张 PNG,固定使用竖向逻辑树布局(主干→彩色分类节点→横线+圆点→内容块),支持表格、象限散点图和要素网格,提供深蓝商务、浅色杂志、科技暗色三套主题,分辨率可调。不适用于可编辑的 XMind/FreeMind 文件、放射状或横向布局、纯文字大纲。
---

# text-to-mindmap

把文本转成一张思维导图 PNG。使用固定的竖向逻辑树设计和 `JSON → HTML → 截图` 渲染流程,保持不同内容的版式一致。

## 这个 Skill 做什么

- 输入:一段内容(已结构化的大纲,或一段原始散文/长文;若来自图片,先把内容读成文本)。
- 中间产物:一份固定 schema 的**结构化 JSON**(见 `references/schema.md`)。
- 输出:一张 **PNG** 思维导图,竖向逻辑树布局,可含表格、象限散点图、要素网格。
- 布局唯一:竖向逻辑树。变化的只是**美术主题**(3 套)和分辨率。

## 何时使用 / 不适用

**用**:用户要把资料、要点、文章、录音纪要等"讲清楚一个主题的内容"变成一张思维导图图片来分享/汇报/发朋友圈。

**不用**:
- 要可再编辑的思维导图源文件(XMind / FreeMind / 幕布)——这是静态 PNG。
- 要放射状、横向左右分支等其它拓扑——本 skill 只做竖向逻辑树。
- 只要纯文字大纲、不需要图。

## 工作流(动手前通读,按序执行)

### Step 1 · 判断输入形态
- **已结构化**(用户给了分类 + 要点):直接进 Step 2 填 JSON。
- **原始散文 / 长文**:先读懂,归纳出"分类 → 每类若干要点"的层级。
- **素材是图片**:先用视觉能力把内容读成文本,再归纳。本 skill 不做"重制原图",只做"把内容做成思维导图"。

### Step 2 · 生成结构化 JSON(按 schema)
- 严格按 `references/schema.md` 的字段写。节点类型:`leaf`(要点)、`table`(对比表)、`chart-quadrant`(象限散点)、`units`(要素/硬件网格)。
- **忠实输入,不外部取数、不杜撰事实**。文中已有的数字/事实照搬;缺关键信息就**回问用户补**,不要自己编,也不要用占位/示意数据。
- 中文文案默认过一遍去 AI 味(调用 `de-ai-flavor-zh` 的准则):要点写成"术语 — 一句人话",别套话、别三段排比、别硬翻译腔。
- 分类数一般 4–8 个;每类要点 3–6 条;能上表格/图的信息优先用 `table`/`chart-quadrant`/`units`,别全堆成文字。
- 分类配色可省略(按主题调色板自动分配),想固定某类颜色再显式给 `color`。

### Step 3 · 选美术主题
- `business-blue`(默认):深色 hero + 白卡片 + 多彩分类节点。通用、商务、投研。
- `magazine-light`:米色亮底 + 衬线标题 + 克制单色。人文、行业观察、有杂志感的分享。
- `tech-dark`:全暗底 + 霓虹连线 + 玻璃卡片。科技、产品发布、年度总结。
- 用户没指定就按主题气质选,并在交付时说明选了哪套、为什么。

### Step 4 · 渲染
```bash
python3 scripts/render.py <你的.json> --out <输出.png> [--theme <主题>] [--width <像素宽>]
```
- `--theme` 缺省时取 JSON 里的 `theme`,再缺省为 `business-blue`。
- 渲染器自动探测本机 Chromium 家族(Chrome/Chromium/Edge/Brave)。没有则报错要求安装——本 skill 必须用真实浏览器渲染,不做替代产出。

### Step 5 · 分辨率处理
- 默认:高清 2x 输出(宽 2360px),锐利不失真。
- 要指定尺寸:`--width 1080`(按宽精确缩放,高等比)。
- 有参考图要对齐它的宽:`--ref-image 参考图.png`(优先级低于 `--width`)。

### Step 6 · 自检闭环(**强制,不可跳过**)
渲染完**用 Read 工具回看输出 PNG**,逐区核对:
1. 每个分类、每条要点都在,没有整块缺失;
2. 没有文字被切一半、表格溢出容器、单元格串行;
3. 中文正常显示,不是方框(方框=缺中文字体);
4. 象限图的点、图例、坐标标签都对得上;主题配色一致不打架。

**不合格就改 JSON 或调整(过长的要点拆短、表格列太多改精简),重渲,再回看。合格才交付。** 交付时说明:主题、尺寸、文件路径。

## Schema 速览(详见 references/schema.md)
```
{ theme, kicker, title, summary, tags[], footer,
  categories: [ { title, color?, items: [
     {type:"leaf",  term, text},
     {type:"table", caption, columns[], widths[], highlightRow?, rows:[{cells:[ "文本" | {meter:"l|m|h",label} ]}]},
     {type:"chart-quadrant", caption, xLabel, yLabel, points:[{label,x(0-1),y(0-1),emphasis?,muted?}]},
     {type:"units", caption, units:[{icon,title,items[]}]}
  ]}]}
```

## 文件
- `assets/template.html` — 逻辑树骨架 + 三套主题 CSS(占位符由脚本注入)。
- `assets/examples/guanghulian.json` — 完整示例,可直接渲染验证:
  `python3 scripts/render.py assets/examples/guanghulian.json --out /tmp/demo.png`
- `scripts/render.py` — 渲染引擎(JSON→HTML→截图→裁剪/缩放)。
- `references/schema.md` — JSON 字段全参考。
- `references/design-spec.md` — 设计系统与主题说明(为什么这么排、怎么改)。
- `references/worked-example.md` — 参考实录:从一张原图到这套设计的三轮迭代与取舍。

## 依赖
- Python `Pillow`(`pip install pillow`)。
- 任一 Chromium 家族浏览器。
- 任一中文字体(PingFang / 思源黑体 / Noto Sans CJK);缺失时脚本会告警。
