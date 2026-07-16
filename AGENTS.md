# my-skill — Codex skills 工作区

这个目录是一批 Codex skill 的打包/编辑工作区。每个子目录是一个独立 skill，结构都是 `SKILL.md`（frontmatter 含 `name` / `description` + 正文），可选 `references/`、`assets/`。

对应的公开仓库：https://github.com/Qmeasure/alex-skill

## skill 清单

| 目录 | 干什么 | 触发场景 | 附属文件 |
|---|---|---|---|
| `de-ai-flavor-zh` | 中文文章去 AI 味、去翻译腔 | "去 AI 味""这段太 AI 了""改得自然点" | `references/案例库.md` |
| `gamma-ppt-outline` | 投研录音转录稿 → Gamma PPT 大纲（.md） | 上传录音转录、"做成 PPT""整理成 Gamma 大纲" | 无 |
| `grill-me` | 拷问式访谈，逐分支压测计划/设计直到达成共识 | "grill me""压测我的方案" | 无 |
| `initial-coverage-advanced` | 机构级中文首次覆盖（深度）研报，五阶段交付 DOCX | "给某公司写首次覆盖/深度研报""覆盖某只股票" | `references/`×7、`assets/`×4、`scripts/`×1 |
| `one-page-model` | 一页式五年损益模型，交付带活公式 Excel | "给 XX 建财务模型/盈利预测""测算盈亏平衡" | `references/`×3 |
| `directional-prompt-writer` | 在方向/方法论层写 prompt（给另一个 AI 执行），只指方向不规定逐段内容 | "帮我写个 prompt""draft a prompt for…" | 无 |
| `verify-before-answer` | 回答前多源交叉验证，杜绝凭记忆作答 | 涉及价格/估值/政策/版本号等时效敏感事实 | 无 |
| `text-to-mindmap` | 文本 → 竖向逻辑树思维导图 PNG（JSON→HTML→截图 确定性管线，3 套美术主题） | "把这段做成思维导图""文本转思维导图""画个思维导图""做成 mindmap" | `assets/template.html`、`assets/examples/`、`scripts/render.py`、`references/`×3 |

## 安装到 Codex

skill 要放进 Codex 的加载路径才生效：全局 `~/.Codex/skills/`，或某项目的 `.Codex/skills/`。本机 `~/.Codex/skills/` 是真实目录，glm / qwen 两个 profile 反向软链到它，装一处三处通用。

装单个 skill（排除 macOS 垃圾文件）：

```bash
rsync -a --exclude '.DS_Store' <skill目录>/ ~/.Codex/skills/<skill名>/
```

装完新开会话才会出现在可调用清单里（skill 列表在会话启动时载入）。

`grill-me` 在本机原本是软链 `~/.Codex/skills/grill-me → ~/.agents/skills/grill-me`；本目录里的是跟进软链复制出来的实体。

## 维护约定

- **改了某个 SKILL.md 后**，如果该 skill 已装进 `~/.Codex/skills/`，记得同步过去（用上面的 rsync），否则线上跑的是旧版。
- **新增 skill**：建子目录 + `SKILL.md`，`name` 用 ASCII 小写连字符且和目录名一致（中文目录名 Codex 不识别）。
- **不提交进公开仓库的东西**（见 `.gitignore`）：`.DS_Store`、`__MACOSX/`、`*.skill`、`*.zip`、`.Codex/`、`AGENTS.md` 本身。
- **公开仓库含第三方来源 skill**（gamma-ppt-outline、strategic-prompt-writer、verify-before-answer、initial-coverage-advanced 等从第三方 `.skill` 包解压而来），改动或再分发时注意来源。

## 推送

仓库 `Qmeasure/alex-skill`，默认分支 `main`。改动可直推 `main`，也可走 PR（`gh pr create` → 合并）。`AGENTS.md` 被 gitignore，推送时不会带上。
