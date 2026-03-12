---
name: graduation-paper-workflow
description: "Run the end-to-end paper workflow from topic discussion to final Word delivery: collect writing constraints, propose topic options, build an outline, expand the outline into markdown draft content, extract formatting requirements from a Word .doc/.docx template into reusable JSON, then combine that JSON with markdown paper text to generate a new Word document that matches the template as closely as possible. Use when Codex needs to help plan, draft, revise, or format a graduation paper or similar coursework paper."
---

# Graduation Paper Workflow

## Overview

Run a multi-stage workflow:

1. Collect writing constraints and required inputs.
2. Propose topic options and help the user choose a title.
3. Build and confirm the paper outline.
4. Expand the outline into a full markdown draft.
5. Extract template structure and formatting into JSON.
6. Combine markdown paper text with the template and JSON to generate a new Word document.

Prefer the documented workflow and bundled scripts first, then patch heuristics and rerun until authoring, extraction, and generation all match the user's requirements and the template.

## Start With Intake

Before drafting, gather the missing inputs progressively. Do not ask everything at once if part of it is already known.

Prioritize these items:

- 专业或所属课程
- 论文方向或想写的技术/设备/场景
- 学校模板文件或最终 Word 格式要求
- 希望的题目风格：偏理论、偏应用、偏案例、偏故障分析、偏发展趋势
- 篇幅或深度要求
- 是否有必须包含或必须避开的内容
- 是否已有参考题目、提纲、初稿或参考资料

If the template is not available yet, still proceed through topic selection, outlining, and markdown drafting, but treat final Word generation as blocked until the template arrives.

## Use The Bundled Scripts

- Run `scripts/extract_word_template_format.py <source.doc|docx> <output.json>` to build or repair the template-format JSON.
- Run `scripts/generate_markdown_papers_docx.py --template-doc <template.doc> --template-json <template.json> --output-dir <dir> <markdown_files...>` to generate `.docx` papers from markdown.

Require a Windows environment with Microsoft Word installed and `pywin32` available, because both scripts use Word COM.

## Follow This Pipeline

1. Collect or confirm the user's constraints: major, direction, template, title preference, scope, and any required content.
2. Generate a topic shortlist and explain each option briefly so the user can choose.
3. After topic confirmation, generate a paper outline with section intent and writing points.
4. After outline confirmation, write a full markdown draft that matches the generator's markdown contract.
5. Inspect the source Word template, any existing format JSON, the markdown input, and the downstream output requirements.
6. Run the extractor and verify page setup, headers, footers, text boxes, semantic blocks, and raw paragraph/run evidence.
7. Patch extractor heuristics if the JSON misclassifies cover items, TOC items, headings, body text, references, or header/footer scope.
8. Once the JSON is trustworthy, run the generator with the template, the JSON, and one or more markdown papers.
9. Validate the generated `.docx` against the template and inspect the output visually in Word when formatting is sensitive.
10. If generation is wrong, patch the generator assumptions or the JSON contract, then rerun.

## Follow These Confirmation Gates

- Confirm the final topic before writing the outline.
- Confirm the outline before expanding into full markdown.
- Confirm the markdown draft before treating it as final input for Word generation.

Keep each gate explicit so the user can redirect the paper before the costly formatting stage.

## Apply These Authoring Rules

- Generate multiple topic candidates before locking the final title.
- For each topic, explain the angle, expected difficulty, and whether the material is easy to fill out into a complete paper.
- Write outlines that are compatible with the markdown contract used by the bundled generator.
- Expand the outline into a complete markdown draft, not just bullet notes.
- Keep the draft close to the selected topic and the confirmed outline; do not silently drift into a different direction.
- If the user supplies only a broad direction, narrow it into concrete, writable topic options before drafting.
- If the user does not provide a template yet, still produce the markdown draft and clearly mark the Word stage as pending template input.

## Apply These Extraction Rules

- Treat `style_name` as weak evidence. Many legacy templates keep everything under one paragraph style and rely on direct formatting instead.
- Inspect text boxes and shapes. Cover titles and other fixed layout elements may live outside the main paragraph flow.
- Distinguish manual TOC from TOC fields. Hand-typed dot leaders and page numbers are plain text, not automatic directory structures.
- Distinguish fill-in lines from headings. Cover metadata lines often use bold + underline + spaces and are not title levels.
- Preserve header/footer scope. Check whether the document uses one section, different first page, or odd/even headers.
- Keep the JSON additive: preserve raw evidence and append higher-level interpretation rather than replacing low-level data.

## Apply These Generation Rules

- Use the template JSON as the formatting contract, not the markdown alone.
- Map markdown content into semantic blocks such as cover title, abstract, keywords, level 1 heading, level 2 heading, body text, conclusion, and references.
- Preserve special layout behavior from the JSON, including text boxes, manual spacing, tabs, fill lines, and line-spacing rules.
- Separate visual formatting from Word navigation levels. Do not copy `outline_level` blindly from extracted paragraph evidence into generated paragraphs.
- Force non-title content such as TOC entries, keywords, body paragraphs, and reference items to body-text outline level so they do not appear as headings in Word navigation.
- Assign navigation levels only to true semantic titles such as abstract title, preface title, level 1 headings, level 2 headings, conclusion title, and references title.
- Keep validation in the loop. If the generated `.docx` diverges from the template, fix the extractor or generator instead of applying one-off manual output edits.

## Read References When Needed

- For the pre-writing flow, required user inputs, and confirmation gates, read [references/authoring-flow.md](references/authoring-flow.md).
- For the full two-stage workflow and failure modes, read [references/workflow.md](references/workflow.md).
- For the expected template JSON contract, read [references/json-contract.md](references/json-contract.md).
- For the current markdown input contract used by the bundled generator, read [references/markdown-contract.md](references/markdown-contract.md).

## Update The Skill

- Add durable workflow rules and failure patterns to `references/workflow.md`.
- Add durable intake, topic-selection, outline, and drafting rules to `references/authoring-flow.md`.
- Add or revise JSON fields in `references/json-contract.md` when the generator starts depending on them.
- Add or revise markdown parsing assumptions in `references/markdown-contract.md` when the input format changes.
- Keep `SKILL.md` short. Move detailed case notes or schema expansions into `references/`.
- If a new fix needs deterministic behavior, patch the bundled scripts and test them before finishing.
