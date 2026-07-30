#!/usr/bin/env python3
"""Run the reusable DOCX-to-editable-PPT pipeline for any strategy report."""

from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import yaml

from scripts.generate_ppt import build_presentation
from scripts.parse_templates import parse_docx
from scripts.plan_presentation_content import build_plan as build_content_plan
from scripts.plan_slides import build_plan as build_slide_plan


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def template_hashes(template_dir: Path) -> dict[str, str]:
    return {
        path.name: sha256(path)
        for path in sorted(template_dir.iterdir())
        if path.is_file()
    }


def require_files(paths: list[Path]) -> None:
    missing = [str(path) for path in paths if not path.is_file()]
    if missing:
        raise FileNotFoundError(
            "Missing required pipeline input(s):\n- " + "\n- ".join(missing)
        )


def count_layout_overlaps(layout_debug: list[dict[str, Any]]) -> int:
    return sum(len(item.get("overlaps") or []) for item in layout_debug)


def run_pipeline(
    project_root: Path,
    input_docx: Path,
    run_dir: Path,
    output_pptx: Path,
    max_pages: int,
) -> dict[str, Any]:
    baseline_output = project_root / "output"
    template_dir = project_root / "template"
    mapping_path = baseline_output / "slide_mapping.yaml"
    layout_path = baseline_output / "ppt_layout.json"
    style_path = baseline_output / "style_config.yaml"
    require_files([input_docx, mapping_path, layout_path, style_path])

    template_before = template_hashes(template_dir)
    run_dir.mkdir(parents=True, exist_ok=True)
    output_pptx.parent.mkdir(parents=True, exist_ok=True)

    # The template mapping, layout and style files are shared read-only
    # configuration. They are not copied or regenerated for each report.
    mapping = yaml.safe_load(mapping_path.read_text(encoding="utf-8"))
    layout = json.loads(layout_path.read_text(encoding="utf-8"))
    style = yaml.safe_load(style_path.read_text(encoding="utf-8"))

    document = parse_docx(input_docx, run_dir)
    doc_payload = {
        "format": "docx",
        "file_count": 1,
        "files": [document],
    }
    doc_structure_path = run_dir / "doc_structure.json"
    write_json(doc_structure_path, doc_payload)

    presentation_plan = build_slide_plan(
        document,
        min_pages=0,
        max_pages=max_pages,
    )
    presentation_plan_path = run_dir / "presentation_plan.json"
    write_json(presentation_plan_path, presentation_plan)

    presentation_content = build_content_plan(
        document,
        presentation_plan["slides"],
    )
    presentation_content_path = run_dir / "presentation_content.json"
    write_json(presentation_content_path, presentation_content)

    generator_inputs = {
        "plan": presentation_plan,
        "content": presentation_content,
        "doc": doc_payload,
        "mapping": mapping,
        "layout": layout,
        "style": style,
        "template_path": (
            project_root
            / "template"
            / style["source_template"]
        ),
        "asset_root": run_dir.parent,
    }
    layout_debug: list[dict[str, Any]] = []
    presentation = build_presentation(generator_inputs, layout_debug)
    presentation.save(output_pptx)
    layout_debug_path = run_dir / "layout_debug.json"
    write_json(
        layout_debug_path,
        {
            "slide_count": len(layout_debug),
            "overlap_count": count_layout_overlaps(layout_debug),
            "slides": layout_debug,
        },
    )

    template_after = template_hashes(template_dir)
    if template_before != template_after:
        raise RuntimeError("Template directory changed during test generation")

    manifest = {
        "run_timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256(input_docx),
        "output_pptx": str(output_pptx),
        "output_pptx_sha256": sha256(output_pptx),
        "planned_slide_count": presentation_plan["planned_slide_count"],
        "generated_slide_count": len(presentation.slides),
        "intermediate_files": {
            "doc_structure": str(doc_structure_path),
            "presentation_plan": str(presentation_plan_path),
            "presentation_content": str(presentation_content_path),
            "slide_mapping": str(mapping_path),
            "ppt_layout": str(layout_path),
            "style_config": str(style_path),
            "layout_debug": str(layout_debug_path),
            "extracted_images": str(run_dir / "docx_images"),
            "extracted_charts": str(run_dir / "docx_charts"),
        },
        "template_files_unchanged": True,
        "template_sha256": template_after,
    }
    write_json(run_dir / "run_manifest.json", manifest)
    return manifest


def main() -> None:
    project_root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=project_root / "test" / "input" / "【广发策略】从杠杆繁荣到筹码松动：韩国杠杆去化走到哪一步？V3(1).docx",
    )
    parser.add_argument(
        "--work-dir",
        "--test-run-dir",
        dest="work_dir",
        type=Path,
        default=project_root / "output" / "work",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=project_root / "test" / "output" / "【广发策略】从杠杆繁荣到筹码松动：韩国杠杆去化走到哪一步？.pptx",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=0,
        help="Optional hard page limit; 0 keeps the dynamically planned count.",
    )
    args = parser.parse_args()

    input_docx = args.input.resolve()
    if not input_docx.is_file():
        raise SystemExit(
            "Test DOCX not found. Place the real report at:\n"
            f"{input_docx}\n"
            "Then run: ./venv/bin/python test_generate.py"
        )

    manifest = run_pipeline(
        project_root=project_root,
        input_docx=input_docx,
        run_dir=args.work_dir.resolve(),
        output_pptx=args.output.resolve(),
        max_pages=args.max_pages,
    )
    print(
        f"Generated {manifest['output_pptx']} "
        f"({manifest['generated_slide_count']} slides)"
    )
    print(f"Intermediate files: {args.work_dir.resolve()}")


if __name__ == "__main__":
    main()
