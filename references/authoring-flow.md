# Authoring Flow

## Goal

Handle the front half of the paper workflow before template extraction and Word generation:

1. Collect user constraints
2. Propose topic options
3. Confirm the final title
4. Build the outline
5. Confirm the outline
6. Expand the outline into a full markdown draft
7. Confirm the markdown draft
8. Hand off to template JSON extraction and final `.docx` generation

## Ask For These Inputs

Collect the missing items progressively. Prefer the smallest set needed to move to the next gate.

Core inputs:

- 专业
- 论文方向或课程方向
- 想写的技术、设备、系统、行业或应用场景
- 学校模板或最终 Word 格式要求

Useful supporting inputs:

- 希望的题目风格：理论分析、应用分析、故障分析、发展趋势、案例研究
- 篇幅或深度要求
- 是否已有参考题目
- 是否有必须包含或必须避开的内容
- 是否已有参考资料、课堂内容、初稿、提纲

## Recommended Intake Order

1. Ask the user's major or course context.
2. Ask what direction or topic family they want to write about.
3. Ask whether a school Word template already exists.
4. Ask whether they prefer a theory-heavy, application-heavy, troubleshooting-heavy, or trend-heavy paper.
5. Ask about any constraints on title, length, or required content.

## Topic Selection Stage

Produce a small candidate list instead of forcing one topic immediately.

For each topic candidate, include:

- Candidate title
- One-sentence angle
- Why it fits the user's major/direction
- Expected writing difficulty
- Whether it is easy to expand into a complete paper

Prefer topics that are:

- Easy to explain with common technical knowledge
- Close to the user's major
- Easy to organize into standard graduation-paper sections
- Compatible with the available template and expected paper length

## Topic Confirmation Gate

Do not start the full outline until the user confirms one of these:

- A chosen title from the candidate list
- A revised title based on the candidate list
- A direct instruction to choose the best title automatically

## Outline Stage

After the title is fixed, generate a structured outline that includes:

- Final title
- Abstract direction
- Keyword suggestions
- Level 1 headings
- Level 2 headings where needed
- What each section should cover
- Reference material direction

Keep the outline easy to expand into the current markdown contract.

## Outline Confirmation Gate

Do not expand into full draft content until the user confirms the outline or gives explicit revision instructions.

Typical revisions at this stage:

- Adjust chapter order
- Add or remove a subsection
- Shift from theory to application
- Narrow or broaden the topic

## Markdown Draft Stage

After outline confirmation, write a complete markdown draft.

The output should be a usable paper draft, not just notes.

It should already conform to the current markdown contract as much as possible:

- Title on the first non-empty line
- `摘要`
- `关键词`
- Body sections with heading patterns the generator can parse
- `六、结束语`
- `参考文献`

## Markdown Confirmation Gate

Treat the markdown draft as the last content checkpoint before formatting.

At this gate, let the user request:

- Tone adjustments
- More detail in specific sections
- Topic refocus
- Simpler or more practical wording
- Reference replacement or cleanup

## Handoff Rule

Only move into template extraction and Word generation when at least one of these is true:

- The user has provided the Word template
- A validated template JSON already exists for the same template
- The user explicitly wants only the markdown draft for now

If the template is missing, finish the markdown stage and clearly say the final Word stage is pending the template.
