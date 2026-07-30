---
name: strategy-slide-planner
description: Dynamically determine slide count from structured Chinese brokerage research, then convert the plan into a traceable Word-to-PPT content layer. Use when planning, compressing, or validating sell-side strategy presentations without generating the final PPT.
---

# Strategy Slide Planner

Create `presentation_plan.json` before `presentation_content.json`. Determine page count from the research rather than inheriting a template deck's slide count. Do not generate or modify a PPT unless the user separately requests it.

## Inputs

Require:

- `doc_structure.json`: title, heading hierarchy, body paragraphs, tables, images, and ordered content.
- `slide_mapping.yaml`: optional compatibility reference for supported template page types and target shape roles.

Preserve the source document's heading order and section logic. Do not introduce facts, causal claims, forecasts, or recommendations absent from the source.

## Workflow

1. Count first-level headings, second-level headings, body characters, tables, and images/charts.
2. Score content importance from conclusions, expectations, inflection points, risks, policy changes, key numbers, and asset-pricing implications.
3. Generate exactly one `cover` and one `summary`.
4. Generate one `section` page for each substantive first-level heading. Treat a risk heading as `risk`, not as a normal section divider.
5. Allocate body pages in source order. Expand dense, important, data-heavy, or visual-heavy subsections; compress short or repetitive subsections.
6. Generate one or two `risk` pages according to risk/disclaimer length and table density.
7. Keep every page type within `cover`, `summary`, `section`, `chart_analysis`, `chart_table`, `table_summary`, `matrix`, and `risk`.
8. Write `presentation_plan.json` with `slide_number`, `section`, `page_type`, `source_heading`, `reason`, and `estimated_content_length`.
9. Use the plan, not the source template page count, to produce `presentation_content.json` for every slide:
   - `slide_number`
   - `title`
   - `key_points`
   - `chart_requirement`
   - `table_requirement`
   - `source`
10. Retain paragraph and table indexes in `source` so every claim remains traceable.
11. Validate page count, field completeness, point count, and source references before delivery.

## Compression Rules

- Write 3–5 core points per analytical page. Exempt `section` pages.
- Put the conclusion or market implication first.
- Preserve key figures, dates, percentage changes, basis points, and forecast intervals.
- Use concise sell-side research language: emphasize expectations, marginal change, catalysts, constraints, and asset-pricing implications.
- Preserve the original argument, qualifications, causal chain, and degree of certainty.
- Do not turn possibilities into certainties or strengthen recommendations.
- Remove repetition, procedural wording, and low-value background before removing evidence.
- Keep each point focused on one claim; use supporting data in the same point when it is essential.

## Page-Type Rules

- `cover`: use the report title and 3–5 research themes; do not add analytical detail.
- `summary`: select 3–5 conclusions spanning the main sections.
- `section`: keep only the section title, an explicit source subtitle when one exists, and the page number. Never generate “本章聚焦”, subsection summaries, body text, or bullets.
- `chart_analysis`: specify the source indicators, comparison dimension, and intended conclusion; use only source data.
- `chart_table`: preserve both the chart thesis and the table's core fields.
- `table_summary`: retain decisive rows and columns plus the conclusion supported by the table.
- `matrix`: preserve row/column semantics, direction symbols, and scenario logic.
- `risk`: preserve risk conditions, transmission mechanisms, and disclaimer boundaries without softening them.

## Output Quality Gate

Reject or revise the plan if:

- slide count differs from `presentation_plan.json`;
- page count was copied from `slide_mapping.yaml` without content-based planning;
- a required field is missing;
- an analytical slide has fewer than 3 or more than 5 points;
- a `section` slide contains generated key points, “本章聚焦” text, body bullets, or other summary copy;
- a key number loses its unit or comparison basis;
- a chart/table requirement lacks a source reference;
- content order or the source's logical qualification changes;
- wording adds unsupported certainty or investment advice.
