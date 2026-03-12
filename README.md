# 毕业论文工作流

`graduation-paper-workflow` 是一个面向 Codex 的开源 skill，用于把课程论文或毕业大作业从前期讨论一路推进到最终 Word 文档交付。

它覆盖的完整流程包括：

1. 收集写作约束和基础信息
2. 生成论文题目候选列表
3. 确认最终题目
4. 生成并确认论文框架
5. 按框架扩写成完整 Markdown 初稿
6. 从 Word 模板中提取格式规则并生成 JSON
7. 使用模板 JSON 和 Markdown 初稿生成最终 `.docx`
8. 校验生成结果

## 仓库内容

- `SKILL.md`
  Skill 主说明文件，供 Codex 直接使用。
- `agents/openai.yaml`
  Skill 的 UI 元数据。
- `references/authoring-flow.md`
  前半段写作流程，包括选题、框架和 Markdown 初稿。
- `references/workflow.md`
  从选题到 Word 生成的整体流程以及常见风险点。
- `references/json-contract.md`
  模板格式 JSON 的结构契约。
- `references/markdown-contract.md`
  生成器要求的 Markdown 输入结构。
- `scripts/extract_word_template_format.py`
  从 `.doc/.docx` 模板提取格式并生成 JSON。
- `scripts/generate_markdown_papers_docx.py`
  根据 Markdown 和模板 JSON 生成最终 Word 文档。

## 适用场景

这个 skill 适合以下类型任务：

- 成人教育毕业大作业
- 专科或本科课程论文
- 需要套学校 Word 模板的论文或报告
- 先写 Markdown，再统一生成 Word 定稿的场景

## 工作流程概览

### 1. 信息收集

优先获取这些内容：

- 专业或课程背景
- 论文方向
- 希望写的技术、设备、系统或应用场景
- 写作风格偏好
- 篇幅或深度要求
- 必须包含或必须避开的内容
- Word 模板文件（如果已经有）

### 2. 题目讨论

先给出一组可选题目，让用户从中选择或微调，而不是一开始就锁死一个题目。

### 3. 框架确认

围绕确认后的题目生成一个可扩写的论文框架，并让用户确认章节结构。

### 4. Markdown 初稿

根据确认后的框架扩写成完整 Markdown 初稿，并保持与生成器的解析约定兼容。

### 5. 模板识别

运行：

```bash
python scripts/extract_word_template_format.py <template.doc|docx> <output.json>
```

### 6. Word 生成

运行：

```bash
python scripts/generate_markdown_papers_docx.py \
  --template-doc <template.doc> \
  --template-json <template.json> \
  --output-dir <output_dir> \
  <paper.md>
```

## 运行要求

- Windows
- Microsoft Word
- Python
- `pywin32`

## 设计要点

- 模板里提取到的 `outline_level` 只当作“证据”，不能无条件照抄到生成结果中。
- 只有真正的语义标题才应该进入 Word 导航。
- 目录条目、关键词、正文段落、参考文献条目必须保持为正文级别。
- 文本框标题等形状对象不能简单按普通段落处理。

## 许可证

本项目使用 MIT License。
