# JSON Contract

## Goal

The template JSON is the contract between extraction and generation.

The extractor should preserve evidence.
The generator should consume the higher-level semantic guidance without losing the raw audit trail.

## Preserve These Top-Level Keys

- `generated_at`
- `source_document`
- `output_purpose`
- `analysis_method`
- `document_stats`
- `page_setup`
- `header_footer_text_snapshot`
- `shape_textboxes_snapshot`
- `style_catalog`
- `format_summary`
- `agent_format_requirements`
- `paragraph_profiles`
- `run_profiles`
- `parts`

## Preserve These Page-Level Facts

Keep `page_setup` authoritative for:

- Page size
- Orientation
- Margins
- Header/footer distances
- `different_first_page_header_footer`
- `odd_and_even_pages_header_footer`

## Preserve These Semantic Guidance Fields

`agent_format_requirements` should carry the higher-level interpretation that downstream generators can use safely:

- `recognition_notes`
- `generation_constraints`
- `layout_notes`
- `recommended_sequence`
- `identified_blocks`

## Preserve These Block Fields

Each item in `identified_blocks` should keep as much of the following as possible:

- `block_id`
- `name`
- `classification`
- `source_kind`
- `paragraph_index` or `shape_index`
- `paragraph_profile_id`
- `run_profile_ids`
- `example_text`
- `resolved_format_snapshot`
- `format_requirements`
- `notes`

Use `source_kind: shape_textbox` for text box content such as cover titles.

## Generation Expectations

The generator can safely depend on these semantics:

- `recommended_sequence` defines the expected block order.
- `identified_blocks` provides at least one concrete example of each block type.
- `format_requirements` carries human-stabilized guidance for fragile layout cases.
- `paragraph_profiles` and `run_profiles` remain available for low-level matching and validation.

## Outline Level Rule

Do not treat extracted `resolved_format_snapshot.paragraph_format.outline_level` as the final navigation contract.

Use this rule instead:

- Treat extracted outline levels as evidence only.
- Let the generator decide Word navigation levels from semantic block type.
- Only true titles should receive non-body outline levels in generated documents.
- TOC entries, keywords, normal body paragraphs, and reference items should stay at body-text outline level even if the source template exposes an outline level on those paragraphs.

## Shape Block Rule

Not every identified block is paragraph-backed.

- Shape-backed blocks such as text-box titles may not have `paragraph_profile_id`.
- Paragraph-driven generators must skip those blocks or handle them via a dedicated shape/text-box output path.

## Preserve Raw Evidence

Do not remove low-level evidence just because a higher-level interpretation now exists.

Keep these sections intact unless the extractor itself changes:

- `parts`
- `paragraph_profiles`
- `run_profiles`

These sections are the audit trail that makes later debugging possible.

## Editing Guidance

- Prefer regenerating the JSON from the extractor after heuristic changes.
- Hand-edit the JSON only when the user explicitly asks for a manual patch or when the extractor cannot yet express the needed rule.
- When adding new semantic keys, do not overwrite unrelated raw fields.
