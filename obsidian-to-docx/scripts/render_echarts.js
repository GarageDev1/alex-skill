#!/usr/bin/env node
/*
 * obsidian-to-docx · ECharts 渲染器
 * 从一个 .md 抽出所有 ```echarts 代码块,逐个用 puppeteer + 系统 Chrome 渲染成 PNG。
 * 输出 chart_01.png, chart_02.png ... 顺序与 md_to_docx.py 的 echarts 计数严格一致。
 *
 * 用法
 *   node render_echarts.js <input.md> <output-dir> [--echarts <echarts.min.js路径或URL>] [--chrome <浏览器路径>]
 *
 * 依赖(按需装其一)
 *   npm ci                 # 安装锁定版本的 puppeteer-core,复用系统浏览器
 *   npm i puppeteer        # 备选:自带 chromium
 *
 * 浏览器路径优先级:--chrome > CHROME_PATH > 常见安装路径 > PATH。
 * ECharts 默认加载本 skill 的 assets/echarts.min.js,无需联网。
 *
 * 设计要点
 * - base point 的 echarts JSON 顶层带 "width"/"height"(900/506);渲染时用它定容器尺寸,
 *   再 strip 掉这两个 key 才 setOption(否则 ECharts 不认 root-level width/height,会告警),
 *   与 xuang.xyz 的 ResearchMarkdown.tsx 处理逻辑一致 → 双端像素级一致。
 * - 图上不画 title.text(base point 约定:H3 当图名);PNG 因此无标题,图名由 docx 端补 caption。
 */

const fs = require("fs");
const path = require("path");

// ---------- 解析参数 ----------
const args = process.argv.slice(2);
if (args.length < 2) {
  console.error("用法: node render_echarts.js <input.md> <output-dir> [--echarts <path|url>] [--chrome <path>]");
  process.exit(1);
}
const inputMd = args[0];
const outDir = args[1];
const defaultEchartsPath = path.resolve(__dirname, "../assets/echarts.min.js");
const echartsArg = argVal("--echarts") || defaultEchartsPath;
const chromePath = argVal("--chrome") || process.env.CHROME_PATH || null;

function argVal(flag) {
  const i = args.indexOf(flag);
  return i >= 0 && i + 1 < args.length ? args[i + 1] : null;
}

// ---------- 从 md 抽 echarts 块 ----------
function extractEcharts(md) {
  const blocks = [];
  const re = /```echarts\s*\n([\s\S]*?)\n```/g;
  let m;
  while ((m = re.exec(md)) !== null) {
    try {
      blocks.push(JSON.parse(m[1]));
    } catch (e) {
      console.warn(`⚠ 第 ${blocks.length + 1} 个 echarts 块 JSON 解析失败,跳过: ${e.message}`);
      blocks.push(null); // 占位,保持编号对齐
    }
  }
  return blocks;
}

// ---------- 载入 puppeteer(core 优先) ----------
function loadPuppeteer() {
  try {
    return { mod: require("puppeteer-core"), core: true };
  } catch (e) {
    try {
      return { mod: require("puppeteer"), core: false };
    } catch (e2) {
      console.error("✗ 未找到 puppeteer-core 或 puppeteer。请先在 skill 根目录执行 `npm ci`。");
      process.exit(1);
    }
  }
}

// 探测 Chrome / Edge / Chromium(显式路径 → 常见安装路径 → PATH)
function findChrome() {
  if (chromePath) {
    if (fs.existsSync(chromePath)) return chromePath;
    console.error(`✗ 指定的浏览器不存在: ${chromePath}`);
    return null;
  }
  const pf = process.env.ProgramFiles || "C:/Program Files";
  const pf86 = process.env["ProgramFiles(x86)"] || "C:/Program Files (x86)";
  const local = process.env.LOCALAPPDATA || "";
  const candidates = [
    path.join(pf, "Google/Chrome/Application/chrome.exe"),
    path.join(pf86, "Google/Chrome/Application/chrome.exe"),
    local && path.join(local, "Google/Chrome/Application/chrome.exe"),
    path.join(pf, "Microsoft/Edge/Application/msedge.exe"),
    path.join(pf86, "Microsoft/Edge/Application/msedge.exe"),
    local && path.join(local, "Microsoft/Edge/Application/msedge.exe"),
    path.join(pf, "Chromium/Application/chrome.exe"),
    path.join(pf86, "Chromium/Application/chrome.exe"),
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
    "/usr/bin/microsoft-edge",
    "/usr/bin/microsoft-edge-stable",
  ];
  for (const c of candidates) {
    if (c && fs.existsSync(c)) return c;
  }
  const names = process.platform === "win32"
    ? ["chrome.exe", "msedge.exe", "chromium.exe"]
    : ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "microsoft-edge", "microsoft-edge-stable"];
  for (const dir of (process.env.PATH || "").split(path.delimiter)) {
    if (!dir) continue;
    for (const name of names) {
      const candidate = path.join(dir, name);
      if (fs.existsSync(candidate)) return candidate;
    }
  }
  return null;
}

async function main() {
  const md = fs.readFileSync(inputMd, "utf-8");
  const options = extractEcharts(md);
  if (!options.length) {
    console.log("无 echarts 块,跳过。");
    return;
  }
  fs.mkdirSync(outDir, { recursive: true });

  const { mod: puppeteer, core } = loadPuppeteer();
  const launchOpts = { headless: "new", args: ["--no-sandbox"] };
  if (core) {
    const chrome = findChrome();
    if (!chrome) {
      console.error("✗ puppeteer-core 需要系统 Chrome/Edge/Chromium,未探测到。用 --chrome <path> 或 CHROME_PATH 指定。");
      process.exit(1);
    }
    launchOpts.executablePath = chrome;
  }

  const browser = await puppeteer.launch(launchOpts);
  const isUrl = /^https?:/.test(echartsArg);
  const echartsPath = isUrl ? null : path.resolve(echartsArg);
  if (echartsPath && !fs.existsSync(echartsPath)) {
    await browser.close();
    throw new Error(`ECharts 脚本不存在: ${echartsPath}`);
  }

  let ok = 0;
  let fail = 0;
  for (let i = 0; i < options.length; i++) {
    const name = `chart_${String(i + 1).padStart(2, "0")}.png`;
    if (!options[i]) continue;
    const outPath = path.join(outDir, name);
    // 断点续跑:已存在的图跳过(改脚本/重跑时不重渲染已成功的)
    if (fs.existsSync(outPath)) { ok++; continue; }
    const w = options[i].width || 900;
    const h = options[i].height || 506;
    // animation:false → echarts 一次性画到最终态,避免截到入场动画中间帧
    const clean = { ...options[i], animation: false };
    delete clean.width;
    delete clean.height;

    // 单图渲染(含 try-catch + 2 次重试, Chrome 资源累积偶发 addScriptTag 失败时不中断整批)
    let done = false;
    for (let attempt = 0; attempt < 2 && !done; attempt++) {
      let page;
      try {
        // 每张图独立 page,避免复用导致的状态污染
        page = await browser.newPage();
        await page.setViewport({ width: w, height: h, deviceScaleFactor: 2 });
        await page.setContent(
          `<!DOCTYPE html><html><head><meta charset="utf-8">` +
            `<style>body{margin:0}#c{width:${w}px;height:${h}px}</style></head>` +
            `<body><div id="c"></div></body></html>`,
          { waitUntil: "domcontentloaded" }
        );
        if (echartsPath) await page.addScriptTag({ path: echartsPath });
        else await page.addScriptTag({ url: echartsArg });

        await page.evaluate(
          (opt, ww, hh) => {
            const ch = window.echarts.init(document.getElementById("c"), null, {
              renderer: "canvas",
              width: ww,
              height: hh,
            });
            ch.setOption(opt);
          },
          clean,
          w,
          h
        );
        await page.waitForFunction("document.querySelector('#c canvas') !== null", { timeout: 15000 });
        await new Promise((r) => setTimeout(r, 300)); // 留一帧给 canvas 栅格化
        const el = await page.$("#c");
        await el.screenshot({ path: outPath });
        done = true;
        ok++;
        console.log(`  ✓ ${name}`);
      } catch (e) {
        if (attempt === 0) { console.log(`  ↻ ${name} 首次失败,重试: ${String(e.message).slice(0, 50)}`); }
        else { console.log(`  ✗ ${name} 渲染失败(跳过,docx 用占位): ${String(e.message).slice(0, 50)}`); fail++; }
      } finally {
        if (page) { try { await page.close(); } catch (_) {} }
      }
    }
  }

  await browser.close();
  console.log(`完成:渲染 ${ok}/${options.length} 张 (失败 ${fail}) → ${outDir}`);
}

main().catch((e) => {
  console.error("✗ 渲染失败:", e);
  process.exit(1);
});
