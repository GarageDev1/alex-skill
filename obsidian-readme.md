# Obsidian 投研工作流

本仓库提供一套可整体迁移的投研工作环境：完成首次覆盖研究、业务驱动量估值、Obsidian Markdown 交付，并可选导出适合打印和对外发送的 Word 文档。

这里的“自包含”是指：研究方法、写作规范、引用文档、建模脚本、校验器、ECharts 资源和 Word 导出工具都在仓库内。公司公告、行情、研报等时效性数据仍需在运行时从公开来源或用户材料取得，并统一整理成 `input.json`，不把个人电脑上的其他 Skill 当作必需依赖。

## 组成与调用关系

```text
用户请求首次覆盖或公司深度拆解
        │
        ▼
equity-research-obsidian（主控与 Obsidian 交付规范）
        ├── initial-coverage-advanced（研究、写作、结构、质检）
        └── valuation-model（业务驱动量模型、Excel、input.json、决策 memo）
                    │
                    ▼
报告 Markdown + 估值 Excel + input.json + 决策 memo
                    │
                    └── obsidian-to-docx（可选：导出 Word）
```

- [`equity-research-obsidian`](./equity-research-obsidian/)：入口 Skill，负责研究管线、Obsidian Markdown、完整 inline ECharts 和跨产物一致性。
- [`initial-coverage-advanced`](./initial-coverage-advanced/)：提供机构级中文研报的研究工序、正文结构、写作规范和质量检查。
- [`valuation-model`](./valuation-model/)：提供业务驱动量估值方法、模型生成、跨平台重算、业务校验、抽表和决策 memo。
- [`obsidian-to-docx`](./obsidian-to-docx/)：可选后处理，将最终 Markdown 和图表导出为 Word，不参与研究与估值判断。

前三个目录必须保持同级，目录名不可更改。ERO 引用同仓的 ICA 与 valuation-model 属于本仓自包含设计的一部分。各 Skill 从自身 `SKILL.md` 所在位置推导仓库根目录，不依赖运行命令时的当前工作目录（CWD）。

## 环境安装

### 1. 估值模型

在 Windows、macOS 或 Linux 上执行：

```text
python -m pip install -r "<repo_root>/valuation-model/requirements.txt"
```

`pywin32` 只会在 Windows 安装；macOS 和 Linux 不需要它。

### 2. Word 导出（可选）

```text
python -m pip install -r "<repo_root>/obsidian-to-docx/requirements.txt"
npm --prefix "<repo_root>/obsidian-to-docx" ci
```

Word 导出默认使用仓库内置的 ECharts 资源，不依赖 CDN。浏览器按以下顺序寻找：显式 `--chrome`、环境变量 `CHROME_PATH`、系统已安装的 Chrome/Edge/Chromium。

### 3. Excel 公式重算后端

`recalc.py --backend auto` 会按当前系统尝试可用后端：

| 系统 | 首选后端 | 后备后端 | 前置条件 |
|---|---|---|---|
| Windows | Microsoft Excel COM | LibreOffice headless | Excel 路径需要 Microsoft Excel 和 `pywin32` |
| macOS | Microsoft Excel AppleScript | LibreOffice headless | Excel 路径需要 Microsoft Excel 和系统自带 `osascript` |
| Linux | LibreOffice headless | — | 安装 LibreOffice，并确保 `soffice` 或 `libreoffice` 在 `PATH` |

自动模式下，前一个后端失败会继续尝试后备后端。所有后端都只重算临时副本；副本通过 XLSX 完整性检查后才原子替换目标文件。没有可用重算后端时，不得进入抽表和报告组装。

## 输出目录

ERO 主流程按以下优先级解析输出目录，命中后停止：

1. 本次调用显式提供的 `output_dir`。
2. 环境变量 `OBSIDIAN_VAULT_ROOT`，输出到 `<OBSIDIAN_VAULT_ROOT>/research/公司研究/<Company>/`。
3. 默认输出到 `<workspace_root>/outputs/equity-research/<Company>/`。

显式相对路径相对于当前 workspace 根目录解析；路径会执行环境变量展开、`~` 展开和绝对化。没有 Obsidian Vault 也能完整生成产物。

如果单独调用 valuation-model，不经过 ERO 主控，则使用：显式输出参数 → `VALUATION_OUTPUT_DIR` → `<workspace_root>/outputs/`。ERO 调用 valuation-model 时会把已经解析好的统一输出目录传下去，避免报告和模型落到不同位置。

四件核心产物必须写入同一目录：

```text
<resolved_output_dir>/
├── <Company>.md
├── <ticker>_valuation_model.xlsx
├── <ticker>_input.json
└── <ticker>_决策memo.md
```

环境变量示例：

```powershell
# Windows PowerShell
$env:OBSIDIAN_VAULT_ROOT = "<your_vault_root>"
$env:VALUATION_OUTPUT_DIR = "<your_output_dir>"
```

```bash
# macOS / Linux
export OBSIDIAN_VAULT_ROOT="<your_vault_root>"
export VALUATION_OUTPUT_DIR="<your_output_dir>"
```

## 路径解析规则

ERO 从自己的位置解析同仓依赖：

```text
ERO_SKILL_ROOT = <repo_root>/equity-research-obsidian
REPO_ROOT       = ERO_SKILL_ROOT.parent
ICA_SKILL_ROOT  = REPO_ROOT/initial-coverage-advanced
VM_SKILL_ROOT   = REPO_ROOT/valuation-model
```

脚本调用必须使用解析后的绝对路径并加引号：

```text
python "<VM_SKILL_ROOT>/scripts/recalc.py" --backend auto "<model.xlsx>"
python "<VM_SKILL_ROOT>/scripts/validate_valuation.py" "<model.xlsx>"
python "<VM_SKILL_ROOT>/scripts/extract_model_tables.py" --full "<model.xlsx>"
```

不要写成依赖 CWD 的 `python valuation-model/scripts/...`，也不要写入个人用户名、桌面目录、固定 Vault 路径或本机代理地址。

## ECharts 与 Obsidian

Obsidian 是预览和知识管理端，不是报告生成前置。

- 报告内包含完整的 ECharts layout，不依赖任何外部自动补全机制。
- 若要在 Obsidian 内交互显示图表，需要安装 `obsidian-echarts` 插件。
- 若只需要 Markdown、Excel 或 Word 交付，可以不安装 Obsidian。
- Word 导出使用 [`obsidian-to-docx/assets`](./obsidian-to-docx/assets/) 中的本地资源，不需要联网加载 ECharts。

## 标准执行顺序

1. 由 ERO 调用 ICA 完成研究、正文和质量检查。
2. 将最新披露、市场数据和研究输入整理为标准 `input.json`。
3. 由 valuation-model 生成业务驱动量估值 Excel。
4. 对模型运行 `recalc.py --backend auto`，回写公式缓存值。
5. 运行 `validate_valuation.py`；只有 `verdict` 不为 `FAIL` 才能继续。
6. 使用 `extract_model_tables.py --full` 从已重算模型抽取报告表格，不手工抄数。
7. 组装最终 Markdown、估值模型、`input.json` 和决策 memo，并检查四者数字一致。
8. 需要打印或对外发送时，再调用 obsidian-to-docx。

数据来源、输入契约和降级顺序见 [`valuation-model/references/02-data-grounding.md`](./valuation-model/references/02-data-grounding.md)。估值章节如何与 ERO 对接，见 [`equity-research-obsidian/references/估值章节数据接入.md`](./equity-research-obsidian/references/估值章节数据接入.md)。

## 交付前检查

- ERO、ICA、valuation-model 仍在同一仓库层级，引用路径没有断裂。
- 输出目录来自显式参数或环境变量，没有个人路径硬编码。
- 模型已重算，且 `validate_valuation.py` 返回 `PASS` 或允许交付的非 FAIL 结果。
- 报告中的估值表来自 `extract_model_tables.py`，不是手工录入。
- ECharts layout 完整，运行与导出均不依赖 CDN。
- Word 导出使用本地资源，并在目标平台完成一次实际渲染检查。

如果 Windows 控制台读取中文文件时出现乱码，可使用 `python -X utf8 ...` 或设置 `PYTHONUTF8=1`；仓库内文本文件统一使用 UTF-8。
