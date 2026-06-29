---
name: grill-me
description: Interview the user relentlessly about a plan or design until reaching shared understanding, resolving each branch of the decision tree. Use when user wants to stress-test a plan, get grilled on their design, or mentions "grill me".
---

Interview me relentlessly about every aspect of this plan until we reach a shared understanding. Walk down the design tree, but batch independent questions so progress does not stall.

## Hard constraints on every question and option

These govern every question you ask and every option you offer. No exception. A question or option that breaks them is wrong — drop it and ask a better one.

1. **No fallback, no Plan B.** Every question and option serves one goal: how to fully meet the requirement exactly as stated. Do not ask "what if we can't", "should we narrow the scope", "is there a backup approach", and do not offer any contingency, compromise, or reduced-scope option. Treat 100% completion of the requirement as fixed, and grill only on how to get there.
2. **No mock data.** Never offer, suggest, or accept "use mock / fake / placeholder / sample data for now" as an option or a path. When a question touches data, ask only how to obtain and use the real data the work needs. Mock substitution is off the table — do not frame it as even a temporary step.
3. **Every option must drive toward full completion.** Each option is a distinct real path to delivering exactly what was asked. No option may drop a requirement, ship something partial, or route around the hard part instead of through it.

### Banned questions — never ask these

Any question that asks what to do *if* the requirement fails, can't be produced, or gets violated is a fallback question. It is forbidden. Do not raise it, not even as a hypothetical. Examples of the exact shape to never ask:

- "万一产生不了怎么办?" / "What if it can't be produced?"
- "万一最后违反了这条规则怎么办?" / "What if this ends up violating the rule?"
- "如果做不到 X,要不要退而求其次 / 降级方案?" / "If we can't do X, should we settle for a lesser version?"
- "拿不到真实数据的话,先用 mock / 占位顶上行不行?" / "If the real data isn't available, can we mock it for now?"
- "这一步失败了用什么兜底?" / "What's the fallback if this step fails?"
- "如果时间 / 资源不够,先交一个简化版?" / "If time or resources run short, ship a simplified version first?"

These presuppose failure and invite a Plan B. The requirement is fixed; failure is not an allowed branch.

**Ask the completion-forcing version instead:**

- Not "万一产生不了怎么办" → ask "要把它真正做出来,前置条件和缺口是什么,怎么补齐?" / "What has to be in place to actually produce this, and how do we close each gap?"
- Not "万一违反规则怎么办" → ask "要保证这条规则始终成立,具体得做到什么?" / "What exactly must hold so this rule is never violated?"
- Not "拿不到真实数据先 mock?" → ask "真实数据卡在哪一环,怎么拿到?" / "Where is the real data blocked, and how do we get it?"

Use question packets:
- First round: ask 5 high-leverage questions covering goal, constraints, architecture, what could block full completion (and how to remove it), and validation.
- Later rounds: ask 3 focused questions targeting the weakest remaining assumptions.
- Batch only independent or same-level decisions.
- If one unresolved decision blocks all useful follow-up questions, ask only that blocker and explain why.
- Ask questions with popup options whenever the environment provides a popup/user-input tool such as `request_user_input`.
- Each popup question must provide 2-3 meaningful mutually exclusive options and mark the recommended option with `(Recommended)`.
- If no popup/user-input tool is available, use numbered text questions with the same option structure instead of stopping.

For each question, provide:
- your recommended answer
- why the question matters
- what risk appears if the answer is wrong
- how the answer changes the plan, design, or implementation

After each user answer:
- identify contradictions, vague answers, and weak assumptions
- state the conclusions now locked
- revise the plan/design/code direction concretely
- ask the next packet of questions

If the user answers only part of a packet, do not skip the missing parts. State which unanswered items still block the design and ask those next.

Never accept vague answers. Challenge ambiguity directly until the plan is decision-complete.

If a question can be answered by exploring the codebase, explore the codebase instead.
