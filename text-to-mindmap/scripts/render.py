#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
text-to-mindmap 渲染引擎:结构化 JSON → 竖向逻辑树思维导图 PNG。

用法:
  python3 render.py mindmap.json --out out.png
  python3 render.py mindmap.json --out out.png --theme tech-dark --width 2360
  python3 render.py mindmap.json --out out.png --ref-image 原图.png   # 对齐参考图宽度

管线:JSON → HTML(套 assets/template.html) → Chromium 无头截图 → PIL 裁掉底部留白
      →(可选)精确缩放到目标宽 → PNG。渲染器自动探测 Chromium 家族。

依赖:Pillow(pip install pillow)、任一 Chromium 内核浏览器、任一中文字体。
JSON schema 见 references/schema.md。
"""
import argparse, html, json, os, shutil, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "assets", "template.html")

THEMES = ("business-blue", "magazine-light", "tech-dark")

# 每套主题的分类配色(JSON 未显式给 color 时按顺序取)
PALETTES = {
    "business-blue": ["#6d5ae6", "#e8943a", "#2fa26a", "#2f7fd4", "#12a3b4", "#8a5cd1", "#20a58c"],
    "magazine-light": ["#c0603a", "#b78b3c", "#4f7a52", "#33628f", "#5b5170", "#8a5a4e", "#3f7a6a"],
    "tech-dark": ["#7c8cff", "#ffb454", "#3ddc97", "#4db5ff", "#28d0e0", "#b48cff", "#3ee6b0"],
}

CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
]
CHROME_WHICH = ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser",
                "microsoft-edge", "brave-browser"]

CJK_HINTS = ("pingfang", "hei", "song", "yahei", "noto", "source han", "sans cjk", "serif cjk")


def find_chrome():
    for p in CHROME_CANDIDATES:
        if os.path.exists(p):
            return p
    for name in CHROME_WHICH:
        p = shutil.which(name)
        if p:
            return p
    return None


def warn_if_no_cjk_font():
    fc = shutil.which("fc-list")
    if not fc:
        return  # 无 fc-list 时不阻断,交给浏览器 font stack 兜底
    try:
        out = subprocess.check_output([fc, ":lang=zh"], text=True, stderr=subprocess.DEVNULL)
        if not out.strip():
            out = subprocess.check_output([fc], text=True, stderr=subprocess.DEVNULL)
            if not any(h in out.lower() for h in CJK_HINTS):
                print("[warn] 未探测到中文字体,渲染出的中文可能是方框。"
                      "请安装 PingFang / 思源黑体 / Noto Sans CJK 之一。", file=sys.stderr)
    except Exception:
        pass


def esc(s):
    return html.escape(str(s), quote=False)


# ---------- 各节点类型 → HTML ----------
def render_leaf(it):
    term = esc(it.get("term", "")).strip()
    text = esc(it.get("text", "")).strip()
    k = f'<span class="k">{term} —</span> ' if term else ""
    return f'<div class="item"><div class="block">{k}{text}</div></div>'


def render_table(it):
    cols = it.get("columns", [])
    widths = it.get("widths", [None] * len(cols))
    ths = "".join(
        f'<th{f" style=\"width:{w}%\"" if w else ""}>{esc(c)}</th>'
        for c, w in zip(cols, widths)
    )
    hl = it.get("highlightRow")
    trs = []
    for ri, row in enumerate(it.get("rows", [])):
        cells = row.get("cells", []) if isinstance(row, dict) else row
        tds = []
        for c in cells:
            if isinstance(c, dict) and "meter" in c:
                lvl = c["meter"]
                lab = esc(c.get("label", ""))
                segs = "".join("<span></span>" for _ in range(3))
                tds.append(f'<td><span class="meter {lvl}"><span class="seg">{segs}</span>'
                           f'<span class="lab">{lab}</span></span></td>')
            else:
                tds.append(f"<td>{esc(c)}</td>")
        cls = ' class="hl"' if hl is not None and ri == hl else ""
        trs.append(f"<tr{cls}>{''.join(tds)}</tr>")
    cap = esc(it.get("caption", ""))
    return (f'<div class="item"><div class="fig"><div class="cap"><i></i>{cap}</div>'
            f'<table><thead><tr>{ths}</tr></thead><tbody>{"".join(trs)}</tbody></table></div></div>')


def render_quadrant(it, color):
    # 坐标 x,y 取值 0..1;原点 (52,222),x 轴到 360,y 轴向上到 22
    def px(x):
        return 52 + max(0.0, min(1.0, x)) * 300
    def py(y):
        return 222 - max(0.0, min(1.0, y)) * 196
    dots, legend = [], []
    for p in it.get("points", []):
        x, y = px(p.get("x", 0.5)), py(p.get("y", 0.5))
        emp = p.get("emphasis", False)
        muted = p.get("muted", False)
        if muted:
            fill, r, tcol, sw = "#c3c8d4", 8, "#8a909d", "#c3c8d4"
        elif emp == "soft":
            fill, r, tcol, sw = color, 10, color, color
            dots.append(f'<circle cx="{x:.0f}" cy="{y:.0f}" r="{r}" fill="{fill}" fill-opacity="0.42"/>')
            dots.append(f'<text x="{x+13:.0f}" y="{y+4:.0f}" font-size="12" fill="{tcol}" font-weight="800" fill-opacity="0.9">{esc(p.get("label",""))}</text>')
            legend.append((f"{color}", "0.42", p.get("label", "")))
            continue
        else:
            fill, r, tcol, sw = color, 10, color, color
        dots.append(f'<circle cx="{x:.0f}" cy="{y:.0f}" r="{r}" fill="{fill}"/>')
        dots.append(f'<text x="{x+13:.0f}" y="{y+4:.0f}" font-size="12" fill="{tcol}" font-weight="{"800" if emp else "400"}">{esc(p.get("label",""))}</text>')
        legend.append((fill, "1", p.get("label", "")))
    leg_html = "".join(
        f'<div class="li"><span class="d" style="background:{c};opacity:{o}"></span>{esc(lbl)}</div>'
        for c, o, lbl in legend
    )
    xlab, ylab = esc(it.get("xLabel", "")), esc(it.get("yLabel", ""))
    cap = esc(it.get("caption", ""))
    svg = f'''<svg viewBox="0 0 380 250" width="100%" style="max-width:470px">
      <defs><marker id="ah" markerWidth="8" markerHeight="8" refX="5" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 Z" fill="#b9bfcc"/></marker></defs>
      <line x1="52" y1="222" x2="52" y2="22" stroke="#d7dbe4" stroke-width="1.5" marker-end="url(#ah)"/>
      <line x1="52" y1="222" x2="360" y2="222" stroke="#d7dbe4" stroke-width="1.5" marker-end="url(#ah)"/>
      <text x="10" y="30" font-size="11" fill="#8a909d">高</text>
      <text x="10" y="220" font-size="11" fill="#8a909d">低</text>
      <text x="60" y="18" font-size="11.5" fill="#565d6b" font-weight="700">{ylab}</text>
      <text x="250" y="242" font-size="11.5" fill="#565d6b" font-weight="700">{xlab}</text>
      <text x="58" y="238" font-size="11" fill="#8a909d">低</text>
      <text x="336" y="220" font-size="11" fill="#8a909d">高</text>
      {''.join(dots)}
    </svg>'''
    return (f'<div class="item"><div class="fig"><div class="cap"><i></i>{cap}</div>'
            f'<div class="quad">{svg}<div class="legend">{leg_html}</div></div></div></div>')


def render_units(it):
    us = it.get("units", [])
    ncol = min(max(len(us), 1), 4)
    boxes = []
    for u in us:
        lis = "".join(f"<li>{esc(x)}</li>" for x in u.get("items", []))
        icon = esc(u.get("icon", "·"))
        boxes.append(f'<div class="unit"><div class="uh"><span class="ic">{icon}</span>'
                     f'{esc(u.get("title",""))}</div><ul>{lis}</ul></div>')
    cap = esc(it.get("caption", ""))
    return (f'<div class="item"><div class="fig"><div class="cap"><i></i>{cap}</div>'
            f'<div class="units" style="--ncol:{ncol}">{"".join(boxes)}</div></div></div>')


def render_category(cat, color):
    items = []
    for it in cat.get("items", []):
        t = it.get("type", "leaf")
        if t == "leaf":
            items.append(render_leaf(it))
        elif t == "table":
            items.append(render_table(it))
        elif t == "chart-quadrant":
            items.append(render_quadrant(it, color))
        elif t == "units":
            items.append(render_units(it))
        else:
            raise ValueError(f"未知节点类型: {t}")
    return (f'<section class="cat" style="--c:{color}">'
            f'<div class="cat-title">{esc(cat.get("title",""))}</div>'
            f'<div class="body">{"".join(items)}</div></section>')


def build_html(data, theme):
    with open(TEMPLATE, encoding="utf-8") as f:
        tpl = f.read()
    palette = PALETTES[theme]
    # hero
    tags = "".join(f"<span>{esc(t)}</span>" for t in data.get("tags", []))
    hero = (f'<div class="hero"><div class="rule"></div>'
            f'<div class="kicker">{esc(data.get("kicker",""))}</div>'
            f'<h1>{esc(data.get("title",""))}</h1>'
            f'<p>{esc(data.get("summary",""))}</p>'
            f'<div class="tags">{tags}</div></div>')
    # tree
    cats = []
    for i, cat in enumerate(data.get("categories", [])):
        color = cat.get("color") or palette[i % len(palette)]
        cats.append(render_category(cat, color))
    foot = ""
    if data.get("footer"):
        foot = f'<div class="foot">{esc(data["footer"])}</div>'
    return (tpl.replace("__THEME__", theme)
               .replace("__HERO__", hero)
               .replace("__TREE__", "\n".join(cats))
               .replace("__FOOT__", foot))


def screenshot(chrome, html_path, raw_png, scale=2, win_w=1180, win_h=8000):
    cmd = [chrome, "--headless", "--disable-gpu", "--hide-scrollbars",
           f"--force-device-scale-factor={scale}", "--default-background-color=FFFFFFFF",
           f"--window-size={win_w},{win_h}", f"--screenshot={raw_png}",
           f"file://{html_path}"]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def crop_and_scale(raw_png, out_png, target_width=None):
    from PIL import Image
    im = Image.open(raw_png).convert("RGB")
    w, h = im.size
    px = im.load()
    bg = px[2, h - 2]

    def has_content(y):
        return any(
            abs(px[x, y][0] - bg[0]) + abs(px[x, y][1] - bg[1]) + abs(px[x, y][2] - bg[2]) > 18
            for x in range(0, w, 6)
        )
    last = h - 1
    while last > 0 and not has_content(last):
        last -= 1
    im = im.crop((0, 0, w, min(h, last + 80)))
    if target_width and target_width != im.size[0]:
        ratio = target_width / im.size[0]
        im = im.resize((target_width, round(im.size[1] * ratio)), Image.LANCZOS)
    im.save(out_png)
    return im.size


def main():
    ap = argparse.ArgumentParser(description="结构化 JSON → 竖向逻辑树思维导图 PNG")
    ap.add_argument("json", help="思维导图结构化 JSON 路径")
    ap.add_argument("--out", required=True, help="输出 PNG 路径")
    ap.add_argument("--theme", default=None, choices=THEMES, help="美术主题(默认取 JSON.theme,否则 business-blue)")
    ap.add_argument("--width", type=int, default=None, help="目标输出像素宽(精确缩放);缺省保持高清 2x")
    ap.add_argument("--ref-image", default=None, help="参考图路径:输出宽度对齐它(优先级低于 --width)")
    ap.add_argument("--scale", type=int, default=2, help="渲染倍率(默认 2)")
    ap.add_argument("--keep-html", action="store_true", help="保留中间 HTML 便于排错")
    args = ap.parse_args()

    with open(args.json, encoding="utf-8") as f:
        data = json.load(f)
    theme = args.theme or data.get("theme") or "business-blue"
    if theme not in THEMES:
        sys.exit(f"未知主题 {theme},可选:{', '.join(THEMES)}")

    chrome = find_chrome()
    if not chrome:
        sys.exit("未找到 Chromium 家族浏览器(Chrome/Chromium/Edge/Brave)。"
                 "请安装其一后重试——本 skill 必须用真实浏览器渲染,不做替代产出。")
    warn_if_no_cjk_font()

    target_w = args.width
    if target_w is None and args.ref_image:
        try:
            from PIL import Image
            target_w = Image.open(args.ref_image).size[0]
        except Exception as e:
            sys.exit(f"读参考图尺寸失败:{e}")

    htmls = build_html(data, theme)
    tmp_html = tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8")
    tmp_html.write(htmls)
    tmp_html.close()
    raw_png = args.out + ".raw.png"
    try:
        screenshot(chrome, os.path.abspath(tmp_html.name), raw_png, scale=args.scale)
        size = crop_and_scale(raw_png, args.out, target_width=target_w)
    finally:
        if os.path.exists(raw_png):
            os.remove(raw_png)
        if args.keep_html:
            final_html = args.out + ".html"
            os.replace(tmp_html.name, final_html)
            print(f"[html] {final_html}")
        elif os.path.exists(tmp_html.name):
            os.remove(tmp_html.name)
    print(f"[ok] theme={theme} size={size[0]}x{size[1]} -> {args.out}")
    print(f"[renderer] {chrome}")


if __name__ == "__main__":
    main()
