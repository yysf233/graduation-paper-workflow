---
name: graduation-paper-workflow
description: "运行从选题讨论到最终 Word 交付的完整论文流程：收集写作约束、生成题目候选、整理并确认框架、扩写成 Markdown 初稿、从 Word 模板提取格式 JSON，并基于模板与 Markdown 生成最终文档。适用于毕业大作业、课程论文及类似写作与排版任务。"
---

# 毕业论文工作流

## 概述

按多阶段流程推进论文任务：

1. 收集写作约束与必要输入
2. 生成题目候选并协助用户选题
3. 生成并确认论文框架
4. 按框架扩写成完整 Markdown 初稿
5. 从 Word 模板中提取结构与格式 JSON
6. 用模板和 Markdown 生成最终 Word 文档

优先使用现有流程和脚本。只有在现有结果不够稳定时，才去修规则或修脚本。

## 先做信息收集

在开始写作前，分阶段收集缺失信息。不要一次把所有问题都抛给用户，优先问当前推进到下一步所必须的信息。

优先确认这些内容：

- 专业或所属课程
- 论文方向，或想写的技术、设备、系统、行业场景
- 学校模板文件，或最终 Word 格式要求
- 题目倾向：理论分析、应用分析、案例分析、故障分析、发展趋势
- 篇幅、深度或难度要求
- 必须包含或必须避开的内容
- 是否已有参考题目、提纲、初稿或资料

如果模板暂时还没有，也可以先完成选题、框架和 Markdown 初稿，只把最终 Word 阶段标记为待模板补齐。

## 使用内置脚本

- 运行 `scripts/extract_word_template_format.py <source.doc|docx> <output.json>` 提取或修复模板格式 JSON。
- 运行 `scripts/generate_markdown_papers_docx.py --template-doc <template.doc> --template-json <template.json> --output-dir <dir> <markdown_files...>` 根据 Markdown 生成 `.docx`。

这两个脚本依赖 Windows、Microsoft Word 和 `pywin32`。

## 按这个流程执行

1. 确认用户约束：专业、方向、模板、题目偏好、篇幅和内容要求。
2. 先生成一个题目候选列表，并简要说明每个题目的角度、难度和可写性。
3. 用户确认题目后，再生成论文框架。
4. 用户确认框架后，再扩写完整 Markdown 初稿。
5. 检查模板、已有 JSON、Markdown 输入和最终输出要求。
6. 运行模板提取器，核对页面设置、页眉页脚、文本框、语义块和底层段落/字符证据。
7. 如果 JSON 对封面、目录、标题、正文、参考文献的识别有系统性偏差，先修提取规则再重跑。
8. 确认 JSON 足够可信后，再运行 Word 生成器。
9. 校验生成的 `.docx`，并在 Word 中人工检查对格式敏感的部分。
10. 如果最终 Word 有问题，优先修生成器规则或 JSON 契约，不要做一次性手工补丁。

## 明确确认关口

- 先确认最终题目，再写框架。
- 先确认框架，再扩写成完整 Markdown。
- 先确认 Markdown 初稿，再进入模板识别和最终 Word 生成。

保持这些关口清晰，避免在高成本排版阶段才发现题目或结构方向错了。

## 写作阶段规则

- 在锁定题目前，先给出多个题目候选。
- 每个题目都要说明写作角度、预期难度、是否容易扩写成完整论文。
- 论文框架必须兼容当前 Markdown 生成器的输入约定。
- Markdown 初稿必须是完整正文，而不是提纲式要点。
- 内容扩写时要严格围绕确认后的题目和框架，不要悄悄偏题。
- 如果用户只给了一个大方向，要先把方向缩小成可写、可成文的具体题目。
- 如果用户暂时没有模板，仍然可以先完成 Markdown 初稿，并明确说明最终 Word 阶段待模板提供后再执行。

## 模板识别规则

- `style_name` 只作为弱证据。很多旧模板整篇都共用一个样式，真正区分结构的是直接格式。
- 必须检查文本框和形状对象。封面主标题等元素可能根本不在正文段落流里。
- 要区分手工目录和自动目录。手打点线和页码不是 TOC 域。
- 要区分封面填写项和真正标题。带下划线和大量空格的封面行通常不是标题层级。
- 要确认页眉页脚范围。检查是否单节、是否首页不同、是否奇偶页不同。
- JSON 要尽量增量保留底层证据，在此基础上再追加高层解释，不要为了“好看”删掉原始信息。

## Word 生成规则

- 生成时以模板 JSON 为排版契约，而不是只看 Markdown 内容。
- 将 Markdown 映射到明确语义块：封面、摘要、关键词、一级标题、二级标题、正文、结束语、参考文献。
- 保留模板中的特殊布局行为，例如文本框、手工留白、Tab、下划线占位和固定行距。
- 必须把“视觉格式”和“Word 导航层级”分开处理，不能直接照搬提取结果中的 `outline_level`。
- 目录条目、关键词、正文段落、参考文献条目必须强制保持正文级别，不能进入 Word 导航。
- 只有真正的语义标题，例如摘要标题、前言标题、一级标题、二级标题、结束语标题、参考文献标题，才应该进入导航。
- 每次生成后都要保留验证环节。如果结果偏了，修规则，不要靠一次性手改文档收尾。

## 需要时再读这些引用文档

- 前置写作流程、用户输入提醒、确认关口： [references/authoring-flow.md](references/authoring-flow.md)
- 完整流程与常见失败模式： [references/workflow.md](references/workflow.md)
- 模板 JSON 契约： [references/json-contract.md](references/json-contract.md)
- Markdown 输入契约： [references/markdown-contract.md](references/markdown-contract.md)

## 维护规则

- 新的写作流程经验、题目讨论规则、框架规则，写进 `references/authoring-flow.md`
- 新的模板识别或生成失败模式，写进 `references/workflow.md`
- 生成器依赖的 JSON 字段变化，写进 `references/json-contract.md`
- Markdown 结构约定变化，写进 `references/markdown-contract.md`
- `SKILL.md` 保持简洁，把长说明沉淀到 `references/`
- 如果某个修复需要稳定重复执行，就改脚本并测试，而不是只改文档说明
