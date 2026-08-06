# 微信公众号 HTML 规则

## 正文文件

正文是一个纯 HTML 片段：

```html
<section style="max-width:100%;box-sizing:border-box;background:#f7f7f5;color:#222222;font-family:inherit;">
  ...
</section>
```

正文只能有一个根 `<section>`，不能包含 `DOCTYPE`、`html`、`head`、`body`。
根元素必须通过 `background` 或 `background-color` 声明可解析的浅色背景。纯色和渐变色标都要保持明亮；不能用透明、继承值或无法判断亮度的写法代替明确背景。示例中的颜色只用于说明格式，不是固定配色。

## 标签和属性

正文可使用：

- 结构：`section`、`p`、`h1`、`h2`、`h3`、`h4`、`blockquote`；
- 列表：`ul`、`ol`、`li`；
- 表格：`table`、`thead`、`tbody`、`tr`、`th`、`td`；
- 行内：`span`、`strong`、`em`、`a`、`br`、`hr`。

限制：

- 结构元素必须有非空 `style` 属性。
- 所有可见文字必须放在 `<span leaf="">` 内。
- `span` 只允许 `leaf` 和 `style`；链接只允许 `href` 和 `style`。
- 表格单元格可以使用 `colspan`、`rowspan` 和 `style`。
- 不使用 `class`、`id`、`data-*`、事件属性和自定义属性。

禁止：

- `img`、`picture`、`source`、SVG、Canvas；
- `video`、`audio`、`iframe`、`object`、`embed`；
- `style`、`script`、`link`；
- `div`、`button`、`form`、`input`；
- `src`、`srcset` 和任何事件属性。

## 内联 CSS

样式既要保证粘贴稳定，也要体现本篇文章的设计方向。字体从公众号编辑器继承；颜色、边框、留白和阴影按文章内容选择，并在全文保持一致。

常用结构规则：

```text
根容器：max-width:100%;box-sizing:border-box;background:<明确的浅色背景>;font-family:inherit
正文：margin:0 0 1em;line-height:1.75;word-break:break-word
标题：margin:1.4em 0 .7em;line-height:1.4;font-weight:700
表格：width:100%;border-collapse:collapse;table-layout:fixed
单元格：padding:.5em;border:1px solid <本篇边框色>;vertical-align:top
链接：text-decoration:underline;word-break:break-all
```

图表可以使用：

- 百分比 `width` 和确定的 `height`；
- `display:block`、`inline-block`、`table`、`table-cell`；
- `vertical-align`、`text-align`；
- `border`、`border-width`、`border-style`、`border-color`；
- 与本篇设计方向一致的颜色、浅色渐变和透明度；
- 克制的 `box-shadow` 或 `text-shadow`；
- `overflow`、`word-break`、`box-sizing`。

根背景可以使用纯色，也可以使用 `background:linear-gradient(...)`、`radial-gradient(...)` 等无需外部资源的原生 CSS 表现。所有色标都要是可解析的浅色。内部标题和强调块可以使用深色背景，校验器只检查根背景的整体明度。

不得使用：

- 外部字体和字体文件；
- CSS 变量；
- `url()`、data URI、`background-image`、`list-style-image`；
- `position:absolute|fixed|sticky`、`float`；
- Grid、动画、媒体查询和关键帧；
- 滤镜；
- 依赖浏览器脚本计算的布局。

## 文字包裹

正确：

```html
<p style="margin:0 0 1em;line-height:1.75;color:inherit;">
  <span leaf="">正文文字</span>
</p>
```

文字不能直接位于 `p`、标题、表格单元格、链接或列表项内。标点也属于可见文字，应放在 `span[leaf]` 中。

## 链接

- 使用完整的 `http://` 或 `https://` 地址。
- 链接文字保留原文，不把 URL 改写成新标题。
- 图片来源只有图片地址时，可以把原题注或替代文字链接到该地址，但不能嵌入图片。
- 禁止 `javascript:`、data URI 和相对路径。

## 预览页

预览页是完整 HTML 文档，可以包含工具栏、按钮、文档级 CSS 和复制脚本。以下条件必须同时满足：

- 正文位于单独的 `gzh-content` 节点；
- 按钮、状态提示和脚本位于正文节点之外；
- 复制函数只选择正文节点的子内容；
- 正文校验器只检查 `{原文件名}_公众号正文.html`，不检查预览外壳。
