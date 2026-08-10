#!/usr/bin/env python3
"""Dynamically plan presentation pages from parsed brokerage research content."""

from __future__ import annotations

import argparse
import json
import math
import re
from pathlib import Path
from typing import Any

from scripts.parse_templates import semantic_heading_level


SUPPORTED_PAGE_TYPES = {
    "cover",
    "summary",
    "section",
    "chart_analysis",
    "chart_table",
    "table_summary",
    "matrix",
    "risk",
    "thanks",
}
IMPORTANCE_TERMS = {
    "核心": 2,
    "结论": 2,
    "拐点": 2,
    "超预期": 2,
    "风险": 2,
    "加息": 1,
    "降息": 1,
    "流动性": 1,
    "资产": 1,
    "通胀": 1,
    "就业": 1,
    "盈利": 1,
    "估值": 1,
    "政策": 1,
}
MAX_VISUALS_PER_SLIDE = 2
SOURCE_LINE_PATTERN = re.compile(r"^(?:数据|资料)来源[:：]", re.I)
POST_VISUAL_CONCLUSION_PATTERN = re.compile(
    r"^(?:因此|这表明|表明|说明|意味着|由此可见|总体来看|整体来看)"
)


def strip_numbering(text: str) -> str:
    return re.sub(
        r"^(?:[一二三四五六七八九十]+、|（[一二三四五六七八九十]+）)\s*",
        "",
        text,
    ).strip()


def item_length(item: dict[str, Any]) -> int:
    if item["type"] == "body":
        return len(re.sub(r"\s+", "", item["text"]))
    if item["type"] == "table":
        return sum(
            len(re.sub(r"\s+", "", cell))
            for row in item.get("rows", [])
            for cell in row
        )
    return 0


def item_visual_count(item: dict[str, Any]) -> int:
    """Return visual capacity consumed by one source item.

    Charts and pictures consume one of two available slots. A genuine editable
    data table consumes the whole slide so its rows and columns remain legible.
    Caption-only wrapper tables without a visual payload consume no slot.
    """

    if item["type"] == "image":
        return 1 if item.get("parent_type") != "table" else 0
    if item["type"] == "chart":
        return len(item.get("chart_indexes", []))
    if item["type"] != "table":
        return 0
    chart_indexes = item.get("chart_indexes", [])
    if chart_indexes:
        return len(chart_indexes)
    image_indexes = item.get("image_indexes", [])
    if image_indexes:
        return len(image_indexes)
    if item.get("row_count", 0) >= 2 and item.get("column_count", 0) >= 2:
        # Compact tables are intentionally allowed beside one chart, matching
        # the template's chart+table analytical layout. Dense tables still own
        # the full page so their rows remain legible.
        return (
            1
            if item.get("row_count", 0) <= 8
            and item.get("column_count", 0) <= 5
            else MAX_VISUALS_PER_SLIDE
        )
    return 0


def is_source_line(item: dict[str, Any]) -> bool:
    return bool(
        item.get("type") == "body"
        and SOURCE_LINE_PATTERN.match(re.sub(r"\s+", " ", item.get("text", "")).strip())
    )


def table_is_matrix(item: dict[str, Any]) -> bool:
    text = " ".join(cell for row in item.get("rows", []) for cell in row)
    matrix_terms = (
        "资产类别",
        "资产价格",
        "价格边际变化",
        "边际变化",
        "推演",
        "情景",
        "方向",
    )
    return sum(term in text for term in matrix_terms) >= 2


def table_is_figure_container(item: dict[str, Any]) -> bool:
    first_text = next(
        (
            cell.strip()
            for row in item.get("rows", [])
            for cell in row
            if cell.strip()
        ),
        "",
    )
    return bool(re.match(r"^图\s*\d+", first_text))


def importance_score(title: str, items: list[dict[str, Any]]) -> int:
    text = title + " " + " ".join(
        item.get("text", "") for item in items if item["type"] == "body"
    )
    score = sum(weight for term, weight in IMPORTANCE_TERMS.items() if term in text)
    numeric_count = len(re.findall(r"\d+(?:\.\d+)?%?|\d+bp", text, flags=re.I))
    return score + min(3, numeric_count // 4)


def parse_sections(document: dict[str, Any]) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    section: dict[str, Any] | None = None
    subsection: dict[str, Any] | None = None
    for item in document["content_order"]:
        kind = item["type"]
        declared_level = (
            item.get("heading_level")
            if item.get("heading_level") is not None
            else 1
            if kind == "heading_1"
            else 2
            if kind == "heading_2"
            else None
        )
        heading_level = semantic_heading_level(
            item.get("text", ""),
            declared_level,
            item.get("style"),
        )
        if heading_level == 1:
            section = {"title": item["text"], "items": [], "subsections": []}
            sections.append(section)
            subsection = None
        elif heading_level == 2 and section is not None:
            subsection = {"title": item["text"], "items": []}
            section["subsections"].append(subsection)
        elif kind in {"body", "table", "image", "chart"} and section is not None:
            normalized = dict(item)
            section["items"].append(normalized)
            if subsection is not None:
                subsection["items"].append(normalized)
    return sections


def distribute_images(sections: list[dict[str, Any]], image_count: int) -> list[int]:
    if image_count <= 0 or not sections:
        return [0] * len(sections)
    weights = [max(1, sum(item_length(item) for item in section["items"])) for section in sections]
    total = sum(weights)
    exact = [image_count * weight / total for weight in weights]
    allocated = [math.floor(value) for value in exact]
    for index in sorted(
        range(len(sections)),
        key=lambda i: exact[i] - allocated[i],
        reverse=True,
    )[: image_count - sum(allocated)]:
        allocated[index] += 1
    return allocated


def split_items(
    title: str,
    items: list[dict[str, Any]],
    image_pressure: int,
) -> list[list[dict[str, Any]]]:
    score = importance_score(title, items)
    target = 950
    if score >= 8:
        target = 700
    elif score >= 5:
        target = 820
    target = max(600, int(target / (1 + min(image_pressure, 2) * 0.12)))

    chunks: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    current_length = 0
    current_visuals = 0
    source_seen_after_visual = False

    def flush() -> None:
        nonlocal current, current_length, current_visuals, source_seen_after_visual
        if current:
            chunks.append(current)
        current = []
        current_length = 0
        current_visuals = 0
        source_seen_after_visual = False

    for item_index, item in enumerate(items):
        length = item_length(item)
        visual_slots = item_visual_count(item)
        source_line = is_source_line(item)
        body_text = re.sub(r"\s+", " ", item.get("text", "")).strip()

        next_visual = next(
            (
                candidate
                for candidate in items[item_index + 1 :]
                if item_visual_count(candidate)
            ),
            None,
        )
        pairs_chart_with_compact_table = bool(
            current_visuals == 1
            and next_visual is not None
            and next_visual["type"] == "table"
            and not next_visual.get("chart_indexes")
            and not next_visual.get("image_indexes")
            and item_visual_count(next_visual) == 1
        )
        starts_new_argument = bool(
            current
            and current_visuals
            and item["type"] == "body"
            and not source_line
            and not pairs_chart_with_compact_table
            and (
                source_seen_after_visual
                or not POST_VISUAL_CONCLUSION_PATTERN.match(body_text)
            )
        )
        exceeds_capacity = bool(
            current and visual_slots and current_visuals + visual_slots > MAX_VISUALS_PER_SLIDE
        )
        splits_long_text_only_unit = bool(
            current
            and not current_visuals
            and not visual_slots
            and current_length + length > target
        )
        should_break = bool(
            starts_new_argument
            or exceeds_capacity
            or splits_long_text_only_unit
        )
        if should_break:
            flush()
        current.append(item)
        current_length += length
        current_visuals += visual_slots
        if source_line and current_visuals:
            source_seen_after_visual = True
    flush()
    chunks = chunks or [[]]

    # Text-length splitting alone is insufficient for chart-heavy reports.
    # Reserve enough pages so content planning can assign no more than two
    # charts/images/tables to each slide. Extra pages intentionally reuse the
    # same source heading; no visible "续" suffix is added.
    visual_count = sum(item_visual_count(item) for item in items)
    required_visual_pages = math.ceil(
        visual_count / MAX_VISUALS_PER_SLIDE
    )
    if len(chunks) < required_visual_pages:
        chunks.extend([[] for _ in range(required_visual_pages - len(chunks))])
    return chunks


def page_type_for(items: list[dict[str, Any]]) -> str:
    tables = [item for item in items if item["type"] == "table"]
    data_tables = [table for table in tables if not table_is_figure_container(table)]
    bodies = [item for item in items if item["type"] == "body"]
    if any(table_is_matrix(table) for table in tables):
        return "matrix"
    if data_tables and bodies:
        return "chart_table"
    if data_tables:
        return "table_summary"
    return "chart_analysis"


def estimated_length(items: list[dict[str, Any]]) -> int:
    return sum(item_length(item) for item in items)


def source_label(section: dict[str, Any], subsection: dict[str, Any] | None) -> str:
    return subsection["title"] if subsection is not None else section["title"]


def reason_for(
    page_type: str,
    items: list[dict[str, Any]],
    score: int,
    part: int,
    parts: int,
    image_pressure: int,
) -> str:
    factors: list[str] = []
    char_count = estimated_length(items)
    if parts > 1:
        factors.append(f"正文信息密度较高，拆分为{parts}页中的第{part}页")
    else:
        factors.append(f"预计承载约{char_count}字")
    table_count = sum(
        item["type"] == "table" and not table_is_figure_container(item)
        for item in items
    )
    figure_count = sum(
        item["type"] == "table" and table_is_figure_container(item)
        for item in items
    )
    if table_count:
        factors.append(f"包含{table_count}个数据表")
    if figure_count:
        factors.append(f"包含{figure_count}个图形容器")
    if image_pressure:
        factors.append(f"本节分配到{image_pressure}个图片/图表资源")
    if score >= 8:
        factors.append("内容重要度高，保留更多论据与关键数据")
    elif score >= 5:
        factors.append("内容重要度中高，适度展开")
    if page_type == "matrix":
        factors.append("表格呈现资产/情景与方向关系，匹配matrix模板")
    return "；".join(factors) + "。"


def add_page(
    pages: list[dict[str, Any]],
    section: str,
    page_type: str,
    source_heading: str,
    reason: str,
    content_length: int,
    source_items: list[dict[str, Any]] | None = None,
) -> None:
    if page_type not in SUPPORTED_PAGE_TYPES:
        raise ValueError(f"Unsupported page type: {page_type}")
    pages.append(
        {
            "slide_number": len(pages) + 1,
            "section": section,
            "page_type": page_type,
            "source_heading": source_heading,
            "reason": reason,
            "estimated_content_length": content_length,
            "source_order_indexes": [
                item["order_index"]
                for item in (source_items or [])
                if item.get("order_index") is not None
            ],
        }
    )


def build_plan(document: dict[str, Any], min_pages: int = 0, max_pages: int = 0) -> dict[str, Any]:
    sections = parse_sections(document)
    if not sections:
        raise ValueError("No first-level sections found in doc_structure.json")

    risk_sections = [
        section
        for section in sections
        if re.match(r"^风险提示(?:及免责声明)?", strip_numbering(section["title"]))
    ]
    substantive_sections = [section for section in sections if section not in risk_sections]
    image_allocations = distribute_images(substantive_sections, len(document.get("images", [])))
    image_allocation_by_section = {
        id(section): allocation
        for section, allocation in zip(substantive_sections, image_allocations)
    }
    pages: list[dict[str, Any]] = []

    add_page(
        pages,
        "封面",
        "cover",
        document["title"],
        "模板约束：每份报告固定生成1页封面。",
        len(document["title"]),
    )
    summary_length = min(
        700,
        sum(estimated_length(section["items"]) for section in substantive_sections) // 8,
    )
    add_page(
        pages,
        "核心结论",
        "summary",
        "全文核心结论",
        "模板约束：每份报告固定生成1页摘要；覆盖主要一级标题的核心结论。",
        summary_length,
    )

    # Every substantive first-level heading owns a section divider, including
    # headings with no second-level headings. "风险提示" is the sole template
    # exception: it maps directly to the inherited risk/disclaimer page rather
    # than producing an otherwise empty numbered divider.
    for section in substantive_sections:
        section_title = strip_numbering(section["title"])
        add_page(
            pages,
            section_title,
            "section",
            section["title"],
            "一级标题对应1页章节页。",
            len(section_title),
        )
        units = section["subsections"] or [
            {"title": section["title"], "items": section["items"]}
        ]
        image_pressure = image_allocation_by_section.get(id(section), 0)
        for unit in units:
            score = importance_score(unit["title"], unit["items"])
            chunks = split_items(unit["title"], unit["items"], image_pressure)
            # Production rule: analytical body pages must contain at least one
            # source chart, image, or table.  Pure-text subsections are already
            # represented in the summary and must not become sparse body pages.
            chunks = [
                chunk
                for chunk in chunks
                if any(item_visual_count(item) for item in chunk)
            ]
            for part, chunk in enumerate(chunks, start=1):
                page_type = page_type_for(chunk)
                add_page(
                    pages,
                    section_title,
                    page_type,
                    source_label(section, unit if section["subsections"] else None),
                    reason_for(
                        page_type,
                        chunk,
                        score,
                        part,
                        len(chunks),
                        image_pressure,
                    ),
                    estimated_length(chunk),
                    chunk,
                )

    risk_items = [item for section in risk_sections for item in section["items"]]
    risk_length = estimated_length(risk_items)
    risk_page_count = 1
    risk_heading = risk_sections[0]["title"] if risk_sections else "风险提示"
    add_page(
        pages,
        "风险提示",
        "risk",
        risk_heading,
        "风险提示不生成编号章节页，直接匹配模板风险提示及免责声明页。",
        risk_length,
    )
    add_page(
        pages,
        "致谢",
        "thanks",
        "模板致谢页",
        "模板约束：每份报告固定生成1页致谢页，并始终置于最终页。",
        0,
    )

    if min_pages and len(pages) < min_pages:
        # Expand the densest analytical pages without introducing new content.
        while len(pages) < min_pages:
            candidates = [
                (index, page)
                for index, page in enumerate(pages)
                if page["page_type"] in {"chart_analysis", "chart_table", "table_summary"}
                and page["estimated_content_length"] >= 500
            ]
            if not candidates:
                break
            index, page = max(
                candidates,
                key=lambda pair: pair[1]["estimated_content_length"],
            )
            half = math.ceil(page["estimated_content_length"] / 2)
            page["estimated_content_length"] = half
            page["reason"] += " 为满足最低页数，对高密度内容进行二次拆分。"
            duplicate = dict(page)
            duplicate["reason"] = "承接上一页的高密度原文内容，不改变原文顺序。"
            pages.insert(index + 1, duplicate)
            for number, item in enumerate(pages, start=1):
                item["slide_number"] = number

    while max_pages > 0 and len(pages) > max_pages:
        merge_candidates: list[tuple[int, int]] = []
        for index in range(2, len(pages) - 1):
            left, right = pages[index], pages[index + 1]
            if (
                left["section"] == right["section"]
                and left["source_heading"] == right["source_heading"]
                and left["page_type"]
                not in {"cover", "summary", "section", "risk", "matrix", "thanks"}
                and right["page_type"]
                not in {"cover", "summary", "section", "risk", "matrix", "thanks"}
            ):
                penalty = (
                    left["estimated_content_length"]
                    + right["estimated_content_length"]
                    + (500 if left["page_type"] != right["page_type"] else 0)
                )
                merge_candidates.append((penalty, index))
        if not merge_candidates:
            raise ValueError(
                f"Cannot compress {len(pages)} planned pages to max_pages={max_pages} "
                "without merging structural, matrix, or risk pages."
            )
        _, index = min(merge_candidates)
        left, right = pages[index], pages[index + 1]
        left["estimated_content_length"] += right["estimated_content_length"]
        left["source_order_indexes"] = list(
            dict.fromkeys(
                [
                    *left.get("source_order_indexes", []),
                    *right.get("source_order_indexes", []),
                ]
            )
        )
        if left["page_type"] != right["page_type"]:
            left["page_type"] = "chart_table"
        left["reason"] += " 为满足最大页数约束，与相邻同主题内容合并。"
        del pages[index + 1]
        for number, item in enumerate(pages, start=1):
            item["slide_number"] = number

    metrics = {
        "heading_1_count": len(sections),
        "heading_2_count": sum(len(section["subsections"]) for section in sections),
        "body_character_count": sum(
            item_length(item)
            for section in sections
            for item in section["items"]
            if item["type"] == "body"
        ),
        "table_count": len(document.get("tables", [])),
        "image_chart_count": len(document.get("images", [])),
        "risk_page_count": risk_page_count,
        "section_page_count": sum(
            page["page_type"] == "section" for page in pages
        ),
    }
    return {
        "source_document": document["file"],
        "planned_slide_count": len(pages),
        "planning_metrics": metrics,
        "template_compatible_page_types": sorted(SUPPORTED_PAGE_TYPES),
        "slides": pages,
    }


def main() -> None:
    project_root = Path(__file__).resolve().parent.parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--doc-structure",
        type=Path,
        default=project_root / "output" / "doc_structure.json",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=project_root / "output" / "presentation_plan.json",
    )
    parser.add_argument("--min-pages", type=int, default=0)
    parser.add_argument(
        "--max-pages",
        type=int,
        default=0,
        help="Optional hard page limit; 0 keeps the dynamically planned count.",
    )
    args = parser.parse_args()

    payload = json.loads(args.doc_structure.read_text(encoding="utf-8"))
    if not payload.get("files"):
        raise SystemExit("doc_structure.json contains no parsed files")
    plan = build_plan(payload["files"][0], args.min_pages, args.max_pages)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(plan, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote {args.output} ({plan['planned_slide_count']} slides)")


if __name__ == "__main__":
    main()
