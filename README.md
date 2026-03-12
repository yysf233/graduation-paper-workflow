# Graduation Paper Workflow

`graduation-paper-workflow` is an open-source Codex skill for taking a coursework or graduation paper from early planning to final Word delivery.

It covers the full flow:

1. Collect writing constraints from the user
2. Propose topic options
3. Confirm the final topic
4. Build and confirm the outline
5. Expand the outline into a full Markdown draft
6. Extract formatting rules from a Word template into JSON
7. Generate a final `.docx` from the template and Markdown draft
8. Validate the generated document

## What This Repository Includes

- `SKILL.md`
  The main skill instructions used by Codex.
- `agents/openai.yaml`
  UI-facing metadata for the skill.
- `references/authoring-flow.md`
  The front-half workflow for topic selection, outlining, and Markdown drafting.
- `references/workflow.md`
  The end-to-end workflow and common failure modes.
- `references/json-contract.md`
  The contract for template-format JSON.
- `references/markdown-contract.md`
  The Markdown structure expected by the bundled generator.
- `scripts/extract_word_template_format.py`
  Extracts formatting information from `.doc/.docx` templates into JSON.
- `scripts/generate_markdown_papers_docx.py`
  Generates `.docx` papers from Markdown plus a template JSON.

## Typical Use Case

This skill is designed for scenarios such as:

- 成人教育毕业大作业
- 专科/本科课程论文
- 需要套学校 Word 模板的论文或报告
- 先写 Markdown，再生成最终 Word 文档的场景

## Workflow Summary

### 1. Intake

Collect the information needed to start:

- Major or course
- Writing direction
- Preferred topic style
- Length or depth expectations
- Must-have and must-avoid content
- Word template, if already available

### 2. Topic Selection

Generate a shortlist of paper topics and let the user choose one.

### 3. Outline Confirmation

Build a clear outline that can be expanded into a complete paper draft.

### 4. Markdown Draft

Write a full draft in Markdown that matches the generator's parsing contract.

### 5. Template Extraction

Run:

```bash
python scripts/extract_word_template_format.py <template.doc|docx> <output.json>
```

### 6. Word Generation

Run:

```bash
python scripts/generate_markdown_papers_docx.py \
  --template-doc <template.doc> \
  --template-json <template.json> \
  --output-dir <output_dir> \
  <paper.md>
```

## Requirements

- Windows
- Microsoft Word
- Python
- `pywin32`

## Design Notes

- The skill treats extracted `outline_level` values as evidence, not as unconditional generation commands.
- Only true semantic headings should appear in Word navigation.
- TOC entries, keywords, normal body paragraphs, and reference items should remain body text in Word navigation.
- Shape-based elements such as text-box titles may need dedicated handling instead of paragraph-only logic.

## License

This repository is released under the MIT License.
