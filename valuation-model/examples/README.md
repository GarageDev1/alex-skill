# examples — worked example + 可复用引擎

> 三个文件、三个角色,按 SKILL.md「沉淀协议」分工:
> - **`build_kit.py`** — 公司无关的可复用引擎(★ 通用逻辑唯一落点, kit-first)
> - **`build_template.py`** — 用 kit 建模的瘦实例(新公司从这里复制起步)
> - **`build_hynix.py`** — SK Hynix v47 真实 worked example(kit 的抽取源; 数字勿照抄, 例≠规范)

## ★ 新公司怎么建模(唯一路径)

1. 复制 `build_template.py`,改全局轴(年份/列/汇率)+ 从标准 `input.json` 填充每个 `write_*` 的 data dict(公开来源、用户材料或可选数据工具均可生产该文件,见 `references/02`);
2. 公司专有的传导公式(分部测算回填、估值对比 lambdas)按自家链写——kit 管结构,链由你接;
3. `PYTHONUTF8=1 python examples/build_<ticker>.py`
4. `python scripts/validate_valuation.py out/<ticker>_valuation_model.xlsx` → **0 fail 才交付**;
5. 运行 `python scripts/recalc.py --backend auto out/<ticker>_valuation_model.xlsx`,再跑一遍 validate;「历史回测≈0」检查有缓存值后才能验。

**验收标准(沉淀协议)**:全程不打开 `build_hynix.py`。任何"不得不去翻它"的瞬间 = kit/reference 欠的债,先还(抽进 kit)再继续。

## kit 提供什么(`build_kit.py`, 对齐 v47 架构)

**13-sheet 标准架构(05 §1, 只向前引用的 DAG)**,每张一个骨架函数:

| 函数 | 产出 sheet | 说明 |
|---|---|---|
| `write_cover` | 封面 | 元信息 + Key Takeaways |
| `write_history` | 历史财务与估值 | 分部+盈利+权益+逐年股本/FX+年末/年均价+实际P/E·P/B+市值+可见备注列+最新季报列(可选) |
| `write_price_chart` | 股价走势 | 月度收盘+阶段+LineChart; 返回 yend/yavg/yhigh/ylow |
| `write_consensus` | 卖方研报共识 | ★按"模型假设"组织的卖方对账单(假设×共识×分歧×我们取的) |
| `write_hist_multiples` | 历史估值倍数 | 数据底座: 自身带+年内高低+同业+相对核心同行比值 |
| `write_multiple_assumptions` | 估值倍数假设 | 为什么主线镜头(业务判断) + 三层分解 × 三案矩阵 + 卖方对账 |
| `write_scenario_switch` | 情景切换 | 下拉开关+嵌套IF案序号+触发条件+杠杆矩阵(组标题/三案故事/历史实际列/当前案INDEX) |
| `write_anchor` | 物理锚 [ANCHOR] | 前瞻列可留 None 由调用方接情景增速 |
| `write_segment_model` | 分部测算 | 铺行+逻辑列, 传导公式调用方回填(每家链不同) |
| `write_fundamentals` | 利润与收入假设 | 假设(段OPM链切换/净利转换/留存, 历史实际打底)+分部营收→段驱动利润→EPS/BPS |
| `write_scenario_valuation` | 情景估值 | 当前案逐年隐含价 + P/E交叉验证 + 当下forward PE证据行 + 市值双币 |
| `write_comparison` | 估值对比 | 三案恒常并排 block(行=调用方 lambdas)+顶部汇总; 强制含 隐含forward P/E 行+回测行 |
| `write_dashboard` | 综合判断仪表盘 | A拐点/B错位GAP(公式)/C催化剂/D情绪 + 综合判断 + 投后跟踪明细 |

**固化的工程坑(helper 层, 调用方不用再操心)**:禁单元格批注(解释走可见列) / 文字开头 "=" 自动转全角 / 案序号嵌套 IF(禁 MATCH+数组常量, WPS兼容) / 逻辑列黑色非斜体+句末分行+去 markdown 粗体 / 行高不锁死 / `finalize()` 年份表头右对齐+冻结窗格。

## 验收基线(2026-06 实测)

- `out/_template_demo.xlsx`(DEMOCO 示意数据):**9 pass / 1 warn / 0 fail**。唯一 warn 是未重算时没有缓存值;重算后估值对比三案回测行历史列应全部接近 0,且全簿公式与禁单元格批注检查通过。
- `out/000660.KS_valuation_model_v47.xlsx`(Hynix 真数据):**8 pass / 2 warn / 0 fail**,新增机检(封面/表序DAG/防污染/每情景隐含forward P/E/禁单元格批注)全过;两个 warn 是既有展示行孤儿提示和未 recalc 缓存值。

## `build_hynix.py` 是什么(诚实版)

SK Hynix (000660.KS) 的 **v47 真实模型生成脚本**(~1170 行, 不依赖 kit——它是 kit 的**抽取源**,kit v2 的全部骨架函数都从它反向剥出来)。它同时是用户在用的活模型,会随季报/研报继续更新。**对其他公司它只有两个用途**:① 当 worked example 看结构;② kit 不够用时按沉淀协议把缺的能力抽进 kit——**禁止直接抄它的数字和判断(例≠规范, 见 references/04 顶部约定)**。

## 已废旧链

旧 `compute_build_pattern.py`(供需取 min + 单元格批注的旧骨架)已删除。新建模只走 `build_kit.py` / `build_template.py`;供需缺口只用于价格路径或可行性旁注,不再作为单公司收入主链。
