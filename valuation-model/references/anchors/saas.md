# 业务驱动量:SaaS(席位 × ACV / 用量计费)

> 适用:订阅软件 / 云 / 用量计费平台。范例:Snowflake、Datadog、Salesforce。

## 基准是什么
两种计费模型,选其一(或混合):
- **席位制**:`付费席位数 × 单席位 ACV`(年合同价)。
- **用量制**:`用量(查询/GB/调用)× 单价`,用量接客户工作负载增长。

## 计算方法(结构)
```
① 业务驱动量:  客户数 × 单客户席位/用量 × 单价
           客户数 = 新增 logo + 存量;净收入留存率 NRR(扩张/流失)是核心乘数
② TAM:     目标客户数 × 成熟 ACV(分客群: 企业/中小)
③ 营收:    ARR = 期初ARR × NRR + 新增ARR ;  营收 ≈ ARR 的期间确认
④ 利润:    毛利(云成本)→ S&M/R&D(获客)→ 经营利润;早期看 Rule of 40 / FCF margin 路径
```

## 数据从哪来 + 怎么估
- **ARR/NRR/客户数/ACV**:公司财报 + 投资者材料(SaaS 披露较全);行业 ACV 基准来自可追溯公开资料、用户材料或可选检索工具。
- **NRR**:接历史 + 同业(>120% 强、<100% 流失),是估值最敏感变量。
- **获客效率**:CAC payback / Magic Number 约束 S&M 投放与盈利路径。
- **估值**:高增长未盈利用 **EV/Sales(配 Rule of 40)+ path to FCF**;成熟后切 DCF / P/E。

## 要点
- **NRR 是杠杆**:留存>100% 时存量自然扩张,模型对 NRR 极敏感,在情景切换里单列 Bear/Base/Bull 三种情景,并写清触发条件。
- ARR(存量复利)和新增 ARR(获客)分开建,不要合成一个收入增速。

## 份额:token 经济 / 私营 frontier lab(用量计费的特例)

> 范例:Anthropic / OpenAI / xAI。公司数据应写入统一 `input.json`，不新增公司专属 Python 脚本。

- **业务驱动量 = token 量 × 有效混合单价**:`GAAP 收入 = token 处理量(Q/yr) × 有效混合单价($/M)`。单价 = 挂牌 mix(Opus/Sonnet/Haiku 加权)× batch/缓存折扣, **商品化逐年下行**(弹性>1, 降价反放大量); token 量校准到卖方收入绝对值, 前瞻 = 上年×(1+情景增速)。改单价或 token 量 → 收入 → 隐含估值全部相关结果随之变化。
- **margin 杠杆 = compute/收入比**(不是 NRR):`EBIT margin ≈ 1 − compute/收入比 − 云分成% − 其他opex%`。compute/收入比能否从高位(60%+)降到稳态(40-50%)是最大不确定, 单列三案。**run-rate(当月×12)≠ GAAP 全年**, 高增长下全年 GAAP ≪ 年末 exit run-rate, 主模型走 GAAP。
- **私营标的 = 整体 EV/$B 口径, 无 per-share**:无公开股价 → 不用 kit 的 `write_history`/`write_fundamentals`(它们绑定 per-share/PE/PB),另行建立"整体 $B"六张表(历史财务/历史倍数/倍数假设/利润/情景估值/估值对比);"当前估值" = 最新融资轮 post-money,"隐含价" = 隐含整体估值($B)并直接与当前估值比较;历史"股价" = 历次融资轮 post-money 估值带,结合当年 run-rate 得到 EV/run-rate de-rating 带。
- **估值方法 = EV/Sales 主 + PE(稳态净利)+ DCF**(亏损刚转正): 主要估值方法逐年 forward 隐含估值 = 收入 × EV/Sales(高增长期高、随增速换挡 normalize); PE 仅 2030E 稳态净利交叉; DCF 做内在价值 sanity check。三种估值方法加权 vs 情景概率加权两口径并呈,分歧 = 对稳态 margin 高位的押注大小。
- **公共实现**:`valuation_kit.methods` 提供 EV/Sales、P/E、P/B、DCF 和 SOTP。每股模型和整体估值模型都应复用这些函数，并由统一输入提供净债务、股本和币种口径。
