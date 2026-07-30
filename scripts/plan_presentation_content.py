#!/usr/bin/env python3
"""Build a Word-to-PPT content plan from parsed DOCX structure and slide mapping."""

from __future__ import annotations

import argparse
import json
import math
import re
from pathlib import Path
from typing import Any

from scripts.parse_templates import semantic_heading_level


ALLOWED_TYPES = {
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
SOURCE_PAGE_TYPES = {
    "chart_analysis",
    "chart_table",
    "table_summary",
    "matrix",
    "risk",
}
SOURCE_PATTERN = re.compile(r"(?:数据|资料)来源[:：]\s*([^\n|]+)", re.I)


def yaml_scalar(value: str) -> Any:
    value = value.strip()
    if value == "null":
        return None
    if value.startswith('"') and value.endswith('"'):
        return value[1:-1]
    if re.fullmatch(r"\d+", value):
        return int(value)
    return value


def load_slide_mapping(path: Path) -> list[dict[str, Any]]:
    """Load the restricted, deterministic YAML schema used by slide_mapping.yaml."""
    slides: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    list_key: str | None = None
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        stripped = raw_line.strip()
        if not stripped or stripped.startswith("#") or stripped == "slides:":
            continue
        if stripped.startswith("- slide_number:"):
            if current is not None:
                slides.append(current)
            current = {"slide_number": int(stripped.split(":", 1)[1].strip())}
            list_key = None
            continue
        if current is None:
            continue
        if stripped.startswith("- ") and list_key:
            current[list_key].append(yaml_scalar(stripped[2:]))
            continue
        key, value = stripped.split(":", 1)
        key, value = key.strip(), value.strip()
        if value:
            current[key] = yaml_scalar(value)
            list_key = None
        else:
            current[key] = []
            list_key = key
    if current is not None:
        slides.append(current)

    expected = list(range(1, len(slides) + 1))
    actual = [slide["slide_number"] for slide in slides]
    if actual != expected:
        raise ValueError(f"Slide numbering must be continuous: {actual}")
    invalid = {slide["page_type"] for slide in slides} - ALLOWED_TYPES
    if invalid:
        raise ValueError(f"Unsupported page types: {sorted(invalid)}")
    return slides


def load_presentation_plan(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    slides = payload.get("slides", [])
    expected = list(range(1, len(slides) + 1))
    actual = [slide.get("slide_number") for slide in slides]
    if actual != expected:
        raise ValueError(f"Presentation plan numbering must be continuous: {actual}")
    invalid = {slide.get("page_type") for slide in slides} - ALLOWED_TYPES
    if invalid:
        raise ValueError(f"Unsupported page types in presentation plan: {sorted(invalid)}")
    return slides


def strip_numbering(text: str) -> str:
    return re.sub(
        r"^(?:[一二三四五六七八九十]+、|（[一二三四五六七八九十]+）)\s*",
        "",
        text,
    ).strip()


def sentence_candidates(text: str) -> list[str]:
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return []
    chunks = re.split(r"(?<=[。！？；])\s*|(?<=；)", text)
    result: list[str] = []
    for chunk in chunks:
        chunk = chunk.strip(" \t\r\n；")
        if len(chunk) < 8:
            continue
        if len(chunk) > 100:
            clauses = [
                clause.strip()
                for clause in re.split(r"[；。]", chunk)
                if len(clause.strip()) >= 8
            ]
            result.extend(clauses or [chunk[:100]])
        else:
            result.append(chunk)
    return result


def point_score(text: str) -> tuple[int, int]:
    conclusion_terms = (
        "意味着",
        "整体来看",
        "核心",
        "因此",
        "预计",
        "可能",
        "风险",
        "拐点",
        "超预期",
        "不变",
        "上调",
        "下调",
    )
    score = sum(3 for term in conclusion_terms if term in text)
    score += min(4, len(re.findall(r"\d+(?:\.\d+)?%?|\d+bp", text, re.I)))
    return score, -len(text)


def excluded_summary_text(text: str) -> bool:
    normalized = re.sub(r"\s+", " ", text).strip()
    return bool(
        not normalized
        or "数据来源" in normalized
        or "资料来源" in normalized
        or re.match(r"^(?:图|表)\s*\d+", normalized)
    )


def clean_research_text(text: str) -> list[str]:
    cleaned: list[str] = []
    for line in re.split(r"[\n|]+", text):
        line = re.sub(r"\s+", " ", line).strip()
        line = SOURCE_PATTERN.sub("", line).strip(" ；;，,")
        if excluded_summary_text(line):
            continue
        cleaned.extend(sentence_candidates(line))
    return cleaned


def conclusion_candidates(items: list[dict[str, Any]]) -> list[str]:
    candidates: list[str] = []
    for item in items:
        if item["type"] == "body":
            candidates.extend(clean_research_text(item.get("text", "")))
        elif item["type"] == "table":
            for row in item.get("rows", []):
                for cell in row:
                    candidates.extend(clean_research_text(cell))
    unique: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        candidate = candidate.rstrip("。") + "。"
        normalized = re.sub(r"[，。；：\s]", "", candidate)
        if normalized and normalized not in seen:
            seen.add(normalized)
            unique.append(candidate)
    return unique


def extract_document_conclusions(
    sections: list[dict[str, Any]],
    minimum: int = 3,
    maximum: int = 5,
) -> list[str]:
    """Select one conclusion-first judgment per chapter, then fill if needed."""
    selected: list[str] = []
    reserve: list[str] = []
    for section in sections:
        candidates = conclusion_candidates(section["items"])
        if not candidates:
            continue
        ranked = sorted(candidates, key=point_score, reverse=True)
        selected.append(ranked[0])
        reserve.extend(ranked[1:])
    for candidate in sorted(reserve, key=point_score, reverse=True):
        if len(selected) >= minimum:
            break
        if candidate not in selected:
            selected.append(candidate)
    return [
        point
        for point in selected[:maximum]
        if not excluded_summary_text(point)
    ]


def compact_points(items: list[dict[str, Any]], minimum: int = 3, maximum: int = 5) -> list[str]:
    candidates: list[tuple[int, str]] = []
    order = 0
    for item in items:
        if item["type"] == "body":
            for sentence in sentence_candidates(item["text"]):
                candidates.append((order, sentence))
                order += 1
        elif item["type"] == "table":
            for row in item.get("rows", [])[1:]:
                cells = [re.sub(r"\s+", " ", cell).strip() for cell in row if cell.strip()]
                if len(cells) >= 2:
                    candidates.append((order, "：".join(cells[:3])))
                    order += 1

    unique: list[tuple[int, str]] = []
    seen: set[str] = set()
    for position, text in candidates:
        normalized = re.sub(r"[，。；：\s]", "", text)
        if normalized and normalized not in seen:
            seen.add(normalized)
            unique.append((position, text.rstrip("。") + "。"))

    if len(unique) <= maximum:
        chosen = unique
    else:
        ranked = sorted(unique, key=lambda item: (point_score(item[1]), -item[0]), reverse=True)
        chosen_positions = {position for position, _ in ranked[:maximum]}
        chosen = [item for item in unique if item[0] in chosen_positions]

    # If a short source unit has fewer than three sentences, split its longest
    # statements conservatively at commas without adding new claims.
    while len(chosen) < minimum and chosen:
        longest_index = max(range(len(chosen)), key=lambda idx: len(chosen[idx][1]))
        position, text = chosen[longest_index]
        clauses = [part.strip("，。 ") for part in text.split("，") if len(part.strip()) >= 8]
        if len(clauses) < 2:
            break
        chosen[longest_index : longest_index + 1] = [
            (position + offset / 10, clause + "。")
            for offset, clause in enumerate(clauses[: minimum - len(chosen) + 1])
        ]
    return [text for _, text in chosen[:maximum]]


def split_section_title(text: str) -> tuple[str, str | None]:
    """Split a section heading only where the source clearly marks a subtitle."""

    normalized = strip_numbering(text)
    if "——" in normalized:
        title, subtitle = normalized.split("——", 1)
        if title.strip() and subtitle.strip():
            return title.strip(), subtitle.strip()
    question_mark = normalized.find("？")
    if 0 <= question_mark < len(normalized) - 1:
        title = normalized[: question_mark + 1].strip()
        subtitle = normalized[question_mark + 1 :].strip()
        if title and subtitle:
            return title, subtitle
    return normalized, None


def normalize_cover_title(text: str) -> str:
    """Preserve the report title while normalizing accidental whitespace."""
    return re.sub(r"[ \t]+", " ", text.replace("\r", "").replace("\n", "")).strip()


def is_risk_section(section: dict[str, Any]) -> bool:
    return bool(re.match(r"^风险提示(?:及免责声明)?", strip_numbering(section["title"])))


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


def first_sentence(text: str) -> str:
    candidates = sentence_candidates(text)
    if candidates:
        return candidates[0].rstrip("。") + "。"
    normalized = re.sub(r"\s+", " ", text).strip()
    return normalized.rstrip("。") + "。" if normalized else ""


def visual_mapping_id(item: dict[str, Any]) -> tuple[str, int]:
    return item["type"], int(item["index"])


def normalized_text_length(text: str) -> int:
    return len(re.sub(r"[\W_]+", "", text, flags=re.UNICODE))


def valid_bold_text(text: str) -> bool:
    return normalized_text_length(text) >= 5


def visual_titles(item: dict[str, Any]) -> list[str]:
    titles: list[str] = []
    texts = [
        str(item.get("caption") or ""),
        str(item.get("text") or ""),
    ]
    texts.extend(
        cell
        for row in item.get("rows", [])
        for cell in row
    )
    for text in texts:
        normalized = re.sub(r"\s+", " ", text).strip()
        if re.match(r"^(?:图|表)\s*\d+\s*[：:]", normalized):
            title = re.sub(
                r"^(?:图|表)\s*\d+\s*[：:]\s*",
                "",
                normalized,
            ).strip()
            if title:
                titles.append(title.rstrip("。") + "。")
    return list(dict.fromkeys(titles))


def semantic_tokens(text: str) -> set[str]:
    normalized = re.sub(
        r"^(?:图|表)\s*\d+\s*[：:]?\s*|(?:数据|资料)来源.*$",
        "",
        text,
        flags=re.I,
    )
    normalized = re.sub(r"[\s，。；：、“”‘’（）()《》？?！!—\-_/]+", "", normalized)
    chinese = "".join(re.findall(r"[\u4e00-\u9fff]", normalized))
    tokens = {
        chinese[index : index + 2]
        for index in range(max(0, len(chinese) - 1))
    }
    tokens.update(
        token.lower()
        for token in re.findall(r"[A-Za-z]+(?:-\d+)?|\d+(?:\.\d+)?%?", text)
    )
    return tokens


def semantic_similarity(left: str, right: str) -> float:
    left_tokens = semantic_tokens(left)
    right_tokens = semantic_tokens(right)
    if not left_tokens or not right_tokens:
        return 0.0
    return len(left_tokens & right_tokens) / math.sqrt(
        len(left_tokens) * len(right_tokens)
    )


def is_summary_sentence(text: str) -> bool:
    return bool(
        re.search(
            r"(?:这表明|表明|说明|意味着|因此|由此可见|总体来看|整体来看|"
            r"核心(?:在于|来源)|主要(?:来自|来源|体现)|尚未|已经进入|仍需|"
            r"风险(?:正在|主要)|边际(?:放缓|改善|趋弱)|当前|当下|相比之下|"
            r"从.+看|反映|可能|仍然|仅(?:回落|出清)|已经|已(?:升至|回落|进入))",
            text,
        )
    )


def paragraph_window(
    items: list[dict[str, Any]],
    visual_position: int,
) -> list[tuple[dict[str, Any], str, int]]:
    window: list[tuple[dict[str, Any], str, int]] = []
    before_distance = 0
    for item in reversed(items[:visual_position]):
        if item["type"] != "body":
            continue
        before_distance += 1
        if before_distance > 3:
            break
        window.append((item, "before", before_distance))
    after_distance = 0
    for item in items[visual_position + 1 :]:
        if item["type"] != "body":
            continue
        after_distance += 1
        if after_distance > 2:
            break
        window.append((item, "after", after_distance))
    return window


def score_candidate(
    candidate_type: str,
    semantic_score: float,
    direction: str,
    distance: int,
) -> dict[str, Any]:
    priority = {
        "visual_title": 4,
        "summary_sentence": 4,
        "semantic_body": 3,
        "explicit_bold": 2,
        "nearest_body": 1,
    }[candidate_type]
    if candidate_type == "visual_title":
        proximity_score = 0.0
    else:
        proximity_score = max(0.0, 8.0 - distance * 2.0)
        if direction == "before":
            proximity_score += 1.0
    total = priority * 100.0 + semantic_score * 50.0 + proximity_score
    return {
        "total": round(total, 4),
        "priority": priority,
        "semantic": round(semantic_score, 4),
        "proximity": round(proximity_score, 4),
    }


def visual_candidates(
    items: list[dict[str, Any]],
    visual_position: int,
) -> list[dict[str, Any]]:
    visual = items[visual_position]
    titles = visual_titles(visual)
    reference_text = " ".join(titles)
    candidates: list[dict[str, Any]] = []

    for title in titles:
        # A title is a high-priority safe fallback, but it should not beat a
        # nearby explanatory conclusion merely by matching itself perfectly.
        score = score_candidate("visual_title", 0.0, "visual", 0)
        candidates.append(
            {
                "paragraph_index": None,
                "paragraph_text": title,
                "mapping_basis": "visual_title",
                "candidate_type": "visual_title",
                "paragraph_order_index": None,
                "direction": "visual",
                "distance": 0,
                "score": score,
            }
        )

    for paragraph, direction, distance in paragraph_window(items, visual_position):
        bold_texts = [
            text.strip().rstrip("。") + "。"
            for text in paragraph.get("bold_sentences", [])
            if valid_bold_text(text)
        ]
        bold_normalized = {
            re.sub(r"[\W_]+", "", text, flags=re.UNICODE)
            for text in bold_texts
        }
        sentences = sentence_candidates(paragraph.get("text", ""))
        for bold_text in bold_texts:
            if bold_text not in sentences:
                sentences.append(bold_text)
        for sentence in sentences:
            sentence = sentence.rstrip("。") + "。"
            normalized_sentence = re.sub(
                r"[\W_]+", "", sentence, flags=re.UNICODE
            )
            semantic_score = semantic_similarity(reference_text, sentence)
            is_bold = any(
                bold
                and (
                    bold in normalized_sentence
                    or normalized_sentence in bold
                )
                for bold in bold_normalized
            )
            if is_summary_sentence(sentence) and (
                semantic_score >= 0.08 or not reference_text
            ):
                candidate_type = "summary_sentence"
            elif semantic_score >= 0.12:
                candidate_type = "semantic_body"
            elif is_bold:
                candidate_type = "explicit_bold"
            else:
                candidate_type = "nearest_body"
            score = score_candidate(
                candidate_type,
                semantic_score,
                direction,
                distance,
            )
            candidates.append(
                {
                    "paragraph_index": paragraph["index"],
                    "paragraph_text": sentence,
                    "mapping_basis": candidate_type,
                    "candidate_type": candidate_type,
                    "paragraph_order_index": paragraph.get("order_index"),
                    "direction": direction,
                    "distance": distance,
                    "score": score,
                }
            )

    unique: dict[tuple[int | None, str, str], dict[str, Any]] = {}
    for candidate in candidates:
        key = (
            candidate["paragraph_index"],
            candidate["paragraph_text"],
            candidate["candidate_type"],
        )
        previous = unique.get(key)
        if previous is None or candidate["score"]["total"] > previous["score"]["total"]:
            unique[key] = candidate
    return sorted(
        unique.values(),
        key=lambda candidate: candidate["score"]["total"],
        reverse=True,
    )


def map_visuals_to_preceding_summaries(
    items: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    mappings: list[dict[str, Any]] = []
    for position, item in enumerate(items):
        is_standalone_image = (
            item["type"] == "image" and item.get("parent_type") != "table"
        )
        has_inline_chart = bool(item.get("chart_indexes"))
        if (
            item["type"] != "table"
            and not is_standalone_image
            and not has_inline_chart
        ):
            continue

        candidates = visual_candidates(items, position)
        candidate = candidates[0] if candidates else {
            "paragraph_index": None,
            "paragraph_text": "",
            "mapping_basis": "missing_candidate",
            "candidate_type": "missing_candidate",
            "paragraph_order_index": None,
            "direction": None,
            "distance": None,
            "score": {
                "total": 0.0,
                "priority": 0,
                "semantic": 0.0,
                "proximity": 0.0,
            },
        }
        visual_type, visual_id = visual_mapping_id(item)
        mappings.append(
            {
                "visual_type": visual_type,
                "visual_id": visual_id,
                "table_id": visual_id if visual_type == "table" else None,
                "image_id": visual_id if visual_type == "image" else None,
                "embedded_image_ids": item.get("image_indexes", []),
                "embedded_chart_ids": item.get("chart_indexes", []),
                "visual_order_index": item.get("order_index"),
                **candidate,
                "candidate_scores": candidates,
            }
        )
    return mappings


def expand_visual_mappings(
    mappings: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Split multi-chart/image Word containers into atomic PPT visuals."""

    expanded: list[dict[str, Any]] = []
    for mapping in mappings:
        chart_ids = list(mapping.get("embedded_chart_ids", []))
        image_ids = list(mapping.get("embedded_image_ids", []))
        if chart_ids:
            for chart_id in chart_ids:
                item = dict(mapping)
                item["embedded_chart_ids"] = [chart_id]
                item["embedded_image_ids"] = []
                expanded.append(item)
        if image_ids:
            for image_id in image_ids:
                item = dict(mapping)
                item["embedded_chart_ids"] = []
                item["embedded_image_ids"] = [image_id]
                expanded.append(item)
        if not chart_ids and not image_ids:
            expanded.append(mapping)
    return expanded


def partition(items: list[Any], count: int) -> list[list[Any]]:
    if count <= 0:
        return []
    if not items:
        return [[] for _ in range(count)]
    return [
        items[math.floor(index * len(items) / count) : math.floor((index + 1) * len(items) / count)]
        for index in range(count)
    ]


def partition_visual_mappings(
    mappings: list[dict[str, Any]],
    count: int,
) -> list[list[dict[str, Any]]]:
    """Keep dense editable tables on their own slide.

    A chart or picture uses one slot; a native table without embedded media
    uses both slots. This preserves the established two-visual maximum while
    preventing a multi-row table from being squeezed beside a chart.
    """

    if count <= 0:
        return []
    groups: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    used_slots = 0
    for mapping in mappings:
        is_native_table = bool(
            mapping.get("table_id") is not None
            and not mapping.get("embedded_chart_ids")
            and not mapping.get("embedded_image_ids")
        )
        slots = 2 if is_native_table else 1
        if current and used_slots + slots > 2:
            groups.append(current)
            current = []
            used_slots = 0
        current.append(mapping)
        used_slots += slots
        if used_slots >= 2:
            groups.append(current)
            current = []
            used_slots = 0
    if current:
        groups.append(current)

    if len(groups) > count:
        raise ValueError(
            f"{len(groups)} visual groups require more than "
            f"{count} planned content slides"
        )
    return groups + [[] for _ in range(count - len(groups))]


def allocate_subsections(subsections: list[dict[str, Any]], page_count: int) -> list[dict[str, Any]]:
    if not subsections:
        return []
    allocations = [1] * len(subsections)
    weights = [max(1, len(section["items"])) for section in subsections]
    while sum(allocations) < page_count:
        index = max(
            range(len(subsections)),
            key=lambda i: weights[i] / allocations[i],
        )
        allocations[index] += 1

    pages: list[dict[str, Any]] = []
    for subsection, count in zip(subsections, allocations):
        for part_index, item_group in enumerate(partition(subsection["items"], count), start=1):
            pages.append(
                {
                    "title": strip_numbering(subsection["title"]),
                    "items": item_group,
                    "part": part_index,
                    "parts": count,
                    "subsection": subsection["title"],
                }
            )
    return pages[:page_count]


def page_units_from_planned_headings(
    section: dict[str, Any],
    content_indexes: list[int],
    mappings: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Build content units from the planner's source-heading decisions."""

    if not section["subsections"]:
        groups = partition(section["items"], len(content_indexes))
        return [
            {
                "title": strip_numbering(section["title"]),
                "items": group,
                "part": part,
                "parts": len(content_indexes),
                "subsection": None,
            }
            for part, group in enumerate(groups, start=1)
        ]

    def heading_key(text: str) -> str:
        return re.sub(r"\s+", "", text)

    subsection_by_key = {
        heading_key(subsection["title"]): subsection
        for subsection in section["subsections"]
    }
    page_counts: dict[str, int] = {}
    for index in content_indexes:
        key = heading_key(str(mappings[index].get("source_heading") or ""))
        page_counts[key] = page_counts.get(key, 0) + 1

    item_groups = {
        key: partition(subsection["items"], page_counts.get(key, 1))
        for key, subsection in subsection_by_key.items()
    }
    cursors: dict[str, int] = {}
    units: list[dict[str, Any]] = []
    for index in content_indexes:
        key = heading_key(str(mappings[index].get("source_heading") or ""))
        subsection = subsection_by_key.get(key)
        if subsection is None:
            raise ValueError(
                f"Planned source heading does not match a subsection: "
                f"{mappings[index].get('source_heading')}"
            )
        part_index = cursors.get(key, 0)
        groups = item_groups[key]
        units.append(
            {
                "title": strip_numbering(subsection["title"]),
                "items": groups[part_index] if part_index < len(groups) else [],
                "part": part_index + 1,
                "parts": len(groups),
                "subsection": subsection["title"],
            }
        )
        cursors[key] = part_index + 1
    return units


def assign_visual_mappings_to_units(
    page_units: list[dict[str, Any]],
    visual_mappings: list[dict[str, Any]],
) -> list[list[dict[str, Any]]]:
    """Assign each visual to its source unit, then spill within that heading."""

    groups: list[list[dict[str, Any]]] = [[] for _ in page_units]
    used_slots = [0 for _ in page_units]

    def unit_owns_mapping(
        unit: dict[str, Any],
        mapping: dict[str, Any],
    ) -> bool:
        for item in unit["items"]:
            if (
                mapping.get("table_id") is not None
                and item["type"] == "table"
                and item.get("index") == mapping["table_id"]
            ):
                return True
            if (
                mapping.get("image_id") is not None
                and item["type"] == "image"
                and item.get("index") == mapping["image_id"]
            ):
                return True
        return False

    # Keep charts/images originating from the same Word table container
    # together. Split only when one container itself carries more than two.
    batches: list[list[dict[str, Any]]] = []
    for mapping in visual_mappings:
        container_key = (
            mapping.get("table_id"),
            mapping.get("visual_order_index"),
        )
        if (
            batches
            and len(batches[-1]) < 2
            and container_key
            == (
                batches[-1][0].get("table_id"),
                batches[-1][0].get("visual_order_index"),
            )
        ):
            batches[-1].append(mapping)
        else:
            batches.append([mapping])

    for batch in batches:
        mapping = batch[0]
        owner = next(
            (
                index
                for index, unit in enumerate(page_units)
                if unit_owns_mapping(unit, mapping)
            ),
            None,
        )
        if owner is None:
            # Chapter-intro visuals can appear before the first subsection and
            # therefore are not present in any subsection item slice. Bind
            # those to the nearest unit by original document order.
            visual_order = mapping.get("visual_order_index")
            ordered_units: list[tuple[int, int]] = []
            if visual_order is not None:
                for index, unit in enumerate(page_units):
                    item_orders = [
                        item.get("order_index")
                        for item in unit["items"]
                        if item.get("order_index") is not None
                    ]
                    if not item_orders:
                        continue
                    distance = min(
                        abs(int(visual_order) - int(order))
                        for order in item_orders
                    )
                    ordered_units.append((distance, index))
            owner = (
                min(ordered_units)[1]
                if ordered_units
                else 0
                if page_units
                else None
            )
        if owner is None:
            raise ValueError("Visual mapping has no available content page")
        is_native_table = bool(
            mapping.get("table_id") is not None
            and not mapping.get("embedded_chart_ids")
            and not mapping.get("embedded_image_ids")
        )
        slots = 2 if is_native_table else len(batch)
        owner_heading = page_units[owner].get("subsection")
        candidate_indexes = [
            index
            for index in range(owner, len(page_units))
            if page_units[index].get("subsection") == owner_heading
        ] + [
            index
            for index in range(0, owner)
            if page_units[index].get("subsection") == owner_heading
        ]
        target = next(
            (
                index
                for index in candidate_indexes
                if used_slots[index] + slots <= 2
            ),
            None,
        )
        if target is None:
            raise ValueError(
                "Planned pages do not provide enough visual capacity for "
                f"{owner_heading or page_units[owner]['title']}"
            )
        groups[target].extend(batch)
        used_slots[target] += slots
    return groups


def requirement(
    mapping: dict[str, Any],
    items: list[dict[str, Any]],
    kind: str,
    visual_mappings: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    page_type = mapping["page_type"]
    if visual_mappings is None:
        table_indexes = [item["index"] for item in items if item["type"] == "table"]
        image_indexes = [
            image_index
            for item in items
            for image_index in item.get("image_indexes", [])
        ]
        chart_indexes = [
            chart_index
            for item in items
            for chart_index in item.get("chart_indexes", [])
        ]
    else:
        table_indexes = [
            item["table_id"]
            for item in visual_mappings
            if item.get("table_id") is not None
        ]
        image_indexes = []
        chart_indexes = []
        for item in visual_mappings:
            if item.get("image_id") is not None:
                image_indexes.append(item["image_id"])
            image_indexes.extend(item.get("embedded_image_ids", []))
            chart_indexes.extend(item.get("embedded_chart_ids", []))
        image_indexes = list(dict.fromkeys(image_indexes))
        chart_indexes = list(dict.fromkeys(chart_indexes))
    if kind == "chart":
        required = page_type in {"chart_analysis", "chart_table"}
        return {
            "required": required,
            "source_image_indexes": image_indexes,
            "source_chart_indexes": chart_indexes,
            "recommended_type": (
                "原文关键指标的趋势或对比图"
                if required
                else None
            ),
            "instruction": (
                "仅使用原文数据，突出拐点、预期差或资产价格方向。"
                if required
                else None
            ),
        }
    required = (
        page_type in {"chart_table", "table_summary", "matrix"}
        or bool(mapping.get("table_shape"))
    )
    return {
        "required": required,
        "source_table_indexes": table_indexes,
        "instruction": (
            "保留原表核心字段与关键数值；矩阵页保持资产类别、方向与推演关系。"
            if required
            else None
        ),
    }


def source_info(
    filename: str,
    section: str | None,
    subsection: str | None,
    items: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "document": filename,
        "section": section,
        "subsection": subsection,
        "paragraph_indexes": [item["index"] for item in items if item["type"] == "body"],
        "table_indexes": [item["index"] for item in items if item["type"] == "table"],
        "image_indexes": [item["index"] for item in items if item["type"] == "image"],
    }


def extract_source_organizations(items: list[dict[str, Any]]) -> list[str]:
    organizations: list[str] = []
    seen: set[str] = set()
    for item in items:
        texts = (
            [item.get("text", "")]
            if item["type"] == "body"
            else [
                *[
                    cell
                    for row in item.get("rows", [])
                    for cell in row
                ],
                str(item.get("source_text") or ""),
            ]
            if item["type"] == "table"
            else []
        )
        for text in texts:
            for match in SOURCE_PATTERN.findall(text):
                for organization in re.split(r"[、,，；;]+", match):
                    organization = re.sub(r"\s+", " ", organization).strip()
                    if not organization or organization in seen:
                        continue
                    seen.add(organization)
                    organizations.append(organization)
    research_center = "广发证券发展研究中心"
    organizations = [
        organization
        for organization in organizations
        if organization != research_center
    ]
    organizations.append(research_center)
    return organizations


def display_source(items: list[dict[str, Any]]) -> str:
    organizations = extract_source_organizations(items)
    if organizations == ["广发证券发展研究中心"]:
        return "资料来源：广发证券发展研究中心"
    return (
        "资料来源："
        + "、".join(organizations[:-1])
        + "，"
        + organizations[-1]
    )


def make_slide(
    mapping: dict[str, Any],
    title: str,
    points: list[str],
    filename: str,
    section: str | None,
    subsection: str | None,
    items: list[dict[str, Any]],
    visual_mappings: list[dict[str, Any]] | None = None,
    rich_text: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    source_trace = source_info(filename, section, subsection, items)
    source_display = (
        display_source(items)
        if mapping["page_type"] in SOURCE_PAGE_TYPES
        else None
    )
    # Compatibility field for the current generator. It contains display-safe
    # credits only; filename and heading paths remain exclusively in source_trace.
    source_compatibility = {
        "document": (
            source_display.removeprefix("资料来源：")
            if source_display
            else ""
        ),
        "section": None,
        "subsection": None,
        "paragraph_indexes": source_trace["paragraph_indexes"],
        "table_indexes": source_trace["table_indexes"],
    }
    return {
        "slide_number": mapping["slide_number"],
        "title": title,
        "key_points": (
            points
            if visual_mappings is not None or rich_text
            else points[:5]
        ),
        "rich_text": rich_text or [],
        "chart_requirement": requirement(
            mapping, items, "chart", visual_mappings=visual_mappings
        ),
        "table_requirement": requirement(
            mapping, items, "table", visual_mappings=visual_mappings
        ),
        "visual_mappings": visual_mappings or [],
        "source_trace": source_trace,
        "display_source": source_display,
        "source": source_compatibility,
    }


def build_plan(document: dict[str, Any], mappings: list[dict[str, Any]]) -> dict[str, Any]:
    sections = parse_sections(document)
    if not sections:
        raise ValueError("No first-level sections found in DOCX structure")
    risk_sections = [section for section in sections if is_risk_section(section)]
    non_risk_sections = [section for section in sections if section not in risk_sections]

    section_slide_indexes = [
        index for index, mapping in enumerate(mappings) if mapping["page_type"] == "section"
    ]
    if len(section_slide_indexes) != len(non_risk_sections):
        raise ValueError(
            "Section-page count in presentation_plan.json must equal "
            "the number of non-risk first-level headings; risk maps directly "
            "to the risk/disclaimer template page"
        )

    slide_units: dict[int, dict[str, Any]] = {}
    boundaries = section_slide_indexes + [
        next(
            (index for index, mapping in enumerate(mappings) if mapping["page_type"] == "risk"),
            len(mappings),
        )
    ]

    for section_number, start_index in enumerate(section_slide_indexes):
        end_index = boundaries[section_number + 1]
        content_indexes = list(range(start_index + 1, end_index))
        section = non_risk_sections[section_number]
        page_units = page_units_from_planned_headings(
            section,
            content_indexes,
            mappings,
        )
        section_visual_mappings = expand_visual_mappings(
            map_visuals_to_preceding_summaries(section["items"])
        )
        visual_groups = assign_visual_mappings_to_units(
            page_units,
            section_visual_mappings,
        )
        for index, unit, visual_group in zip(
            content_indexes,
            page_units,
            visual_groups,
        ):
            unit["section"] = section["title"]
            unit["visual_mappings"] = visual_group
            slide_units[index] = unit

    risk_indexes = [
        index for index, mapping in enumerate(mappings) if mapping["page_type"] == "risk"
    ]
    risk_section = risk_sections[0] if risk_sections else {
        "title": "风险提示",
        "items": [],
        "subsections": [],
    }
    # Risk is a fixed text-only template page. Keep only the leading risk
    # paragraphs and stop when the report enters appendix/author/legal tables.
    # Those trailing tables and images must never consume visual capacity or
    # create additional risk slides.
    risk_text_items: list[dict[str, Any]] = []
    for item in risk_section["items"]:
        if item["type"] != "body":
            break
        risk_text_items.append(item)
    risk_item_groups = partition(risk_text_items, len(risk_indexes))
    risk_visual_groups = [[] for _ in risk_indexes]
    for index, group, visual_group in zip(
        risk_indexes,
        risk_item_groups,
        risk_visual_groups,
    ):
        slide_units[index] = {
            "title": "风险提示及免责声明",
            "items": group,
            "part": 1,
            "parts": 1,
            "section": risk_section["title"],
            "subsection": None,
            "visual_mappings": visual_group,
        }

    slides: list[dict[str, Any]] = []
    filename = document["file"]
    for index, mapping in enumerate(mappings):
        page_type = mapping["page_type"]
        if page_type == "cover":
            overview_items = [
                item for section in non_risk_sections for item in section["items"]
            ]
            points = [
                f"研究主线：{strip_numbering(section['title'])}。"
                for section in non_risk_sections[:5]
            ]
            slides.append(
                make_slide(
                    mapping,
                    normalize_cover_title(document["title"]),
                    points,
                    filename,
                    None,
                    None,
                    overview_items,
                )
            )
        elif page_type == "summary":
            source_items: list[dict[str, Any]] = []
            for section in non_risk_sections:
                source_items.extend(section["items"])
            metadata = document.get("report_metadata", {})
            summary_paragraphs = metadata.get("summary_paragraphs") or []
            if summary_paragraphs:
                points = [paragraph["text"] for paragraph in summary_paragraphs]
            else:
                report_summary = (metadata.get("summary") or "").split(
                    "风险提示", 1
                )[0]
                points = clean_research_text(report_summary)
                if not points:
                    points = extract_document_conclusions(
                        non_risk_sections,
                        minimum=3,
                        maximum=5,
                    )
            slides.append(
                make_slide(
                    mapping,
                    "核心结论",
                    points,
                    filename,
                    None,
                    None,
                    source_items,
                    rich_text=summary_paragraphs,
                )
            )
        elif page_type == "section":
            section_number = section_slide_indexes.index(index)
            section = non_risk_sections[section_number]
            section_title, section_subtitle = split_section_title(section["title"])
            section_slide = make_slide(
                mapping,
                section_title,
                [],
                filename,
                section["title"],
                None,
                section["items"],
            )
            section_slide["subtitle"] = section_subtitle
            slides.append(section_slide)
        elif page_type == "thanks":
            slides.append(
                make_slide(
                    mapping,
                    "",
                    [],
                    filename,
                    None,
                    None,
                    [],
                )
            )
        else:
            unit = slide_units[index]
            title = unit["title"]
            visual_mappings = unit.get("visual_mappings", [])
            is_disclaimer = (
                page_type == "risk"
                and risk_indexes
                and index != risk_indexes[0]
            )
            if page_type == "risk":
                visual_mappings = []
                if is_disclaimer:
                    title = "免责声明"
                    disclaimer = (
                        document.get("report_metadata", {}).get(
                            "legal_disclaimer"
                        )
                        or "免责声明内容以原始研究报告为准。"
                    )
                    points = [disclaimer]
                else:
                    title = "风险提示"
                    points = [
                        item["text"].rstrip("；;。") + "。"
                        for item in risk_text_items
                        if item["type"] == "body" and item.get("text")
                    ][:5]
            else:
                points = [
                    visual_mapping["paragraph_text"]
                    for visual_mapping in visual_mappings
                    if visual_mapping["paragraph_text"]
                ]
                if not points:
                    points = compact_points(
                        unit["items"],
                        minimum=1,
                        maximum=3,
                    )
            planned_slide = make_slide(
                mapping,
                title,
                points,
                filename,
                unit["section"],
                unit["subsection"],
                unit["items"],
                visual_mappings,
            )
            planned_slide["content_mode"] = (
                "disclaimer" if is_disclaimer else "bullets"
            )
            slides.append(planned_slide)

    return {
        "source_document": filename,
        "slide_count": len(slides),
        "slides": slides,
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
        "--presentation-plan",
        type=Path,
        default=project_root / "output" / "presentation_plan.json",
    )
    parser.add_argument(
        "--slide-mapping",
        type=Path,
        default=project_root / "output" / "slide_mapping.yaml",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=project_root / "output" / "presentation_content.json",
    )
    parser.add_argument(
        "--debug-output",
        type=Path,
        default=project_root / "output" / "content_mapping_debug.json",
    )
    args = parser.parse_args()

    doc_payload = json.loads(args.doc_structure.read_text(encoding="utf-8"))
    if not doc_payload.get("files"):
        raise SystemExit("doc_structure.json contains no parsed files")
    document = doc_payload["files"][0]
    mappings = (
        load_presentation_plan(args.presentation_plan)
        if args.presentation_plan.exists()
        else load_slide_mapping(args.slide_mapping)
    )
    plan = build_plan(document, mappings)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(plan, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    debug_mappings = [
        {
            "slide_number": slide["slide_number"],
            "visual_type": mapping["visual_type"],
            "table_id": mapping["table_id"],
            "image_id": mapping["image_id"],
            "embedded_image_ids": mapping["embedded_image_ids"],
            "embedded_chart_ids": mapping["embedded_chart_ids"],
            "paragraph_index": mapping["paragraph_index"],
            "paragraph_text": mapping["paragraph_text"],
            "mapping_basis": mapping["mapping_basis"],
            "candidate_type": mapping["candidate_type"],
            "direction": mapping["direction"],
            "distance": mapping["distance"],
            "score": mapping["score"],
            "candidate_scores": mapping["candidate_scores"],
        }
        for slide in plan["slides"]
        for mapping in slide.get("visual_mappings", [])
    ]
    args.debug_output.parent.mkdir(parents=True, exist_ok=True)
    args.debug_output.write_text(
        json.dumps(
            {
                "mapping_count": len(debug_mappings),
                "mappings": debug_mappings,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"Wrote {args.output}")
    print(f"Wrote {args.debug_output}")


if __name__ == "__main__":
    main()
