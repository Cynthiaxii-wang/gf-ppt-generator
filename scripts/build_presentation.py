"""
Build PPT structure from parsed Word document.

Input:
    output/test_run/doc_structure.json

Output:
    output/test_run/presentation_plan.json
    output/test_run/presentation_content.json
    output/test_run/visual_mapping.json

Current goal:
    doc_structure.json -> PPT page structure
"""

import json
from pathlib import Path


def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )


def clean_title(text):
    """
    Remove report numbering.
    """

    if not text:
        return ""

    prefixes = [
        "一、",
        "二、",
        "三、",
        "四、",
        "五、",
        "六、",
        "七、",
        "八、",
        "九、",
        "十、",
        "（一）",
        "（二）",
        "（三）",
        "（四）",
        "（五）",
        "（六）",
    ]

    for prefix in prefixes:
        text = text.replace(prefix, "")

    return text.strip()



def get_body_after_heading(body, start_index, end_index):
    """
    Extract paragraphs between two headings.
    """

    paragraphs = []

    for item in body:

        idx = item.get("index")

        if idx is None:
            continue

        if start_index < idx < end_index:

            # only normal paragraphs
            if item.get("heading_level") is None:

                text = item.get("text", "").strip()

                # filter short noise
                if len(text) > 20:
                    paragraphs.append(text)

    return paragraphs



def build_slides(report):

    slides = []

    headings1 = report.get("heading_1", [])
    body = report.get("body", [])


    # cover

    slides.append(
        {
            "page": 1,
            "page_type": "cover",
            "title": report.get(
                "title",
                "研究报告"
            ),
            "content": []
        }
    )


    # summary

    metadata = report.get(
        "report_metadata",
        {}
    )

    slides.append(
        {
            "page": 2,
            "page_type": "summary",
            "title": "核心观点",
            "content": [
                metadata.get(
                    "summary",
                    ""
                )
            ]
        }
    )


    page = 3


    for i, heading in enumerate(headings1):

        title = clean_title(
            heading.get(
                "text",
                ""
            )
        )


        # section page

        slides.append(
            {
                "page": page,
                "page_type": "section",
                "title": title,
                "content": []
            }
        )

        page += 1


        start = heading.get(
            "index",
            0
        )


        if i + 1 < len(headings1):

            end = headings1[i + 1].get(
                "index",
                999999
            )

        else:

            end = 999999


        paragraphs = get_body_after_heading(
            body,
            start,
            end
        )


        current = []


        for paragraph in paragraphs:

            current.append(
                paragraph
            )

            # 每页最多3段

            if len(current) >= 3:

                slides.append(
                    {
                        "page": page,
                        "page_type": "content",
                        "title": title,
                        "content": current
                    }
                )

                page += 1
                current = []


        if current:

            slides.append(
                {
                    "page": page,
                    "page_type": "content",
                    "title": title,
                    "content": current
                }
            )

            page += 1


    return slides




def build_visual_mapping(report):

    mappings = []

    body = report.get(
        "body",
        []
    )


    for item in body:

        text = item.get(
            "text",
            ""
        )


        if text.startswith("图"):

            mappings.append(
                {
                    "source_index": item.get(
                        "index"
                    ),
                    "title": text,
                    "type": "image"
                }
            )


        elif text.startswith("表"):

            mappings.append(
                {
                    "source_index": item.get(
                        "index"
                    ),
                    "title": text,
                    "type": "table"
                }
            )


    return mappings




def main():

    project_root = Path(__file__).resolve().parent.parent


    input_file = (
        project_root
        /
        "output"
        /
        "test_run"
        /
        "doc_structure.json"
    )


    output_dir = (
        project_root
        /
        "output"
        /
        "test_run"
    )


    doc = load_json(
        input_file
    )


    # ⭐关键修复：
    # doc_structure 的真实结构
    # 顶层 -> files -> 第一篇报告

    report = doc["files"][0]


    slides = build_slides(
        report
    )


    visual_mapping = build_visual_mapping(
        report
    )


    save_json(
        {
            "slides": slides
        },
        output_dir /
        "presentation_plan.json"
    )


    save_json(
        {
            "slides": slides
        },
        output_dir /
        "presentation_content.json"
    )


    save_json(
        visual_mapping,
        output_dir /
        "visual_mapping.json"
    )


    print(
        f"Generated {len(slides)} slides"
    )



if __name__ == "__main__":
    main()