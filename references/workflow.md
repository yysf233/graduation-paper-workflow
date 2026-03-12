# Workflow

## Goal

Run the full pipeline:

1. `Collect constraints -> topic list`
2. `Confirmed topic -> confirmed outline`
3. `Confirmed outline -> markdown draft`
4. `Word template -> format JSON`
5. `Markdown paper text + template + format JSON -> new Word document`

## Prerequisites

- Run on Windows.
- Require Microsoft Word to be installed.
- Require Python with `pywin32`.
- Prefer working from the original `.doc` or `.docx`, not only from an older JSON.

## Stage 0: Collect Constraints

1. Confirm the user's major or course context.
2. Confirm the intended paper direction.
3. Ask whether the Word template already exists.
4. Ask about desired style, scope, and any must-have or must-avoid content.

If the template is not available, continue through topic, outline, and markdown stages, but keep the final Word stage pending.

## Stage 1: Propose And Confirm Topic

1. Generate a candidate title list.
2. Explain why each option fits.
3. Let the user choose or revise the title.
4. Do not start the full outline until the topic is confirmed.

## Stage 2: Build And Confirm Outline

1. Generate a structured outline for the selected topic.
2. Include abstract direction, keywords, section hierarchy, and writing points.
3. Confirm the outline before drafting full markdown.

## Stage 3: Write And Confirm Markdown Draft

1. Expand the confirmed outline into full markdown.
2. Keep the markdown aligned with the parser contract used by the generator.
3. Let the user request revisions.
4. Treat the revised markdown as the final content input for formatting.

## Stage 4: Extract Template JSON

1. Inspect the source template and the current JSON, if one exists.
2. Run `scripts/extract_word_template_format.py`.
3. Open the generated JSON and inspect these areas first:
   - `page_setup`
   - `header_footer_text_snapshot`
   - `shape_textboxes_snapshot`
   - `agent_format_requirements`
4. Compare the JSON against the actual Word document in Word.
5. Patch the extractor when recognition is systematically wrong.
6. Rerun the extractor and verify the changed fields.

## Stage 5: Generate Word From Markdown

1. Confirm the template JSON is trustworthy before using it for generation.
2. Read the markdown paper and check that it matches the generator's expected section markers and heading structure.
3. Run `scripts/generate_markdown_papers_docx.py`.
4. Inspect the generated `.docx` in Word.
5. Review the generated `validation_report.json` if the script emits one.
6. Patch generator assumptions or the JSON contract when formatting or classification is wrong.

## High-Risk Extraction Failures

- Everything shares the same `style_name`, so style-based heading detection fails.
- Cover titles or logos live in text boxes, not in body paragraphs.
- The TOC is manual text, not a TOC field.
- Cover metadata uses underline + spaces as fill lines and gets misread as headings.
- Keywords lines use mixed run formatting, where only part of the line is bold.
- Body paragraphs use tabs or leading spaces instead of clean paragraph indentation.
- Headers and footers appear on all pages because the file has a single section with no first-page exception.

## High-Risk Generation Failures

- The drafting stage produces markdown that does not match the current parser contract.
- The selected topic drifts between title, outline, and final markdown draft.
- Markdown headings do not map cleanly to the template's semantic blocks.
- Cover title, metadata lines, or TOC get written as normal paragraphs even when the template expects text boxes or manual layout.
- Reference items are converted into automatic numbering even though the template uses plain text.
- The JSON preserves raw evidence, but the generator ignores `format_requirements` and uses only profile IDs.
- The generator copies extracted `outline_level` directly, causing body text, keywords, TOC items, or reference items to appear as headings in Word navigation.
- Shape-based identified blocks such as text-box titles do not have `paragraph_profile_id`, so a paragraph-only generator can crash if it assumes every block has one.
- Validation passes structurally but the visual result still drifts because of tabs, text boxes, or mixed run formatting.

## Preferred Fix Strategy

- Fix the intake or topic-selection stage when the paper direction itself is weak or unstable.
- Fix the outline stage when the markdown draft is hard to expand or does not match the intended structure.
- Fix the markdown draft before touching formatting when the content itself is still wrong.
- Fix the extractor when the issue is template interpretation.
- Fix the generator when the issue is markdown-to-block mapping or output layout.
- Preserve raw extracted evidence in `parts`, `paragraph_profiles`, and `run_profiles`.
- Add semantic interpretation in `agent_format_requirements`.
- Treat navigation outline as generator-owned semantics. Use extracted `outline_level` only as evidence, not as an unconditional generation command.
- Skip or specially handle shape-based blocks in paragraph-driven generation code paths.
- Keep validation in the loop after every deterministic fix.

## Current Proven Pattern From The Graduation Template

Use this pattern as a known-good example when a school paper template behaves similarly:

- The user often needs help choosing a title before any formatting work can begin.
- A confirmed outline makes the markdown draft much more stable and reduces later rework.
- The cover main title is a text box.
- Most paragraphs still share one base style, so style names are not enough.
- The TOC is hand-typed with dot leaders.
- The file uses one section, with no different first-page header/footer.
- Cover, abstract, body, and references are separated mainly by direct formatting, spacing, tabs, and layout position.
- Extracted `outline_level` values are noisy and not safe to reuse directly for generation.
- Correct generation requires forcing正文类内容 back to body-text outline level, even if the template evidence shows an outline level on those paragraphs.

## Extend Later

When the user continues to add steps, append only durable workflow rules here:

- New extraction heuristics
- New intake prompts and topic-selection rules
- New outline or drafting rules
- New markdown parsing assumptions
- New validation steps
- New JSON fields that downstream generators depend on
- New handling rules for tables, shapes, numbering, or section breaks
