# Markdown Contract

## Goal

Define the markdown structure expected by the current bundled generator.

This is the current contract, not a universal markdown standard. Update it when the generator changes.

This contract also defines what the authoring stage must produce before handing content to the final Word generation stage.

## Current Expected Structure

- The first non-empty line is the paper title.
- The markdown includes these section markers:
  - `摘要`
  - `关键词`
  - `参考文献`
- The current generator also treats `六、结束语` as the start of the conclusion section.

## Current Parsing Assumptions

- Lines between `摘要` and `关键词` are abstract paragraphs.
- The first non-empty line after `关键词` is the keywords content.
- Body lines after the keywords section are classified by text pattern:
  - `^[一二三四五六七八九十]+、` -> level 1 heading
  - `^（[一二三四五六七八九十]+）` -> level 2 heading
  - other non-empty lines -> body paragraphs
- Lines after `六、结束语` go to the conclusion block.
- Lines after `参考文献` become reference items.

## Authoring Output Rule

When the skill writes the markdown draft from topic discussion and outline:

- Write the title directly on the first non-empty line.
- Emit the exact section markers expected here.
- Keep the heading hierarchy compatible with the generator's regex rules.
- Prefer stable plain-text headings over decorative markdown patterns that the generator does not parse.
- Treat the markdown draft as generator input, not only as a human-readable note.

## Navigation Mapping Rule

The current generator should map markdown content to Word navigation levels by semantic role, not by copied template outline evidence:

- Abstract title, preface title, conclusion title, and references title -> navigation title level
- `^[一二三四五六七八九十]+、` headings -> level 1 heading in navigation
- `^（[一二三四五六七八九十]+）` headings -> level 2 heading in navigation
- Keywords, body paragraphs, TOC entries, and reference items -> body text, not navigation headings

## Current Reference Assumption

- Reference lines may be written as `[1] ...`
- The current generator rewrites those lines to `1 ...`
- If the markdown already contains plain numbered references, the generator keeps them as text

## When To Update This Contract

Update this file when any of these happen:

- The markdown heading scheme changes
- The conclusion marker changes
- The keywords format changes
- The reference format changes
- The generator begins to support richer markdown constructs such as lists, tables, or images
