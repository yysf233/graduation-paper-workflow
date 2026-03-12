from __future__ import annotations

import argparse
import gc
import json
import os
import re
import shutil
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import win32com.client

from extract_word_template_format import extract_document_dict


WD_PAGE_BREAK = 7
WD_INFO_ADJUSTED_PAGE_NUMBER = 1
# Keep visual paragraph formatting separate from Word navigation semantics.
# Only true semantic titles should receive non-body outline levels.
WD_OUTLINE_LEVEL_1 = 1
WD_OUTLINE_LEVEL_2 = 2
WD_OUTLINE_LEVEL_BODY_TEXT = 10

ALIGNMENT_MAP = {
    "left": 0,
    "center": 1,
    "right": 2,
    "both": 3,
    "justify": 3,
}

LINE_SPACING_RULE_MAP = {
    "auto": 1,
    "exact": 4,
}

UNDERLINE_MAP = {
    None: 0,
    "single": 1,
}

FONT_NAME_MAP = {
    "SimSun": "宋体",
    "SimHei": "黑体",
    "FangSong_GB2312": "FangSong_GB2312",
    "KaiTi": "楷体",
}

FONT_ALIAS_MAP = {
    "SimSun": "simsun",
    "宋体": "simsun",
    "SimHei": "simhei",
    "黑体": "simhei",
    "FangSong_GB2312": "fangsong_gb2312",
    "仿宋_GB2312": "fangsong_gb2312",
    "仿宋": "fangsong",
    "KaiTi": "kaiti",
    "楷体": "kaiti",
}

WORD_SIMSUN_SUBSTITUTE_BLOCKS = {
    "toc_title",
    "toc_entry",
    "preface_title",
    "body_text",
    "conclusion_title",
}

H1_RE = re.compile(r"^[一二三四五六七八九十]+、")
H2_RE = re.compile(r"^（[一二三四五六七八九十]+）")
REF_RE = re.compile(r"^\[(\d+)\]\s*(.*)$")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

for prefix, uri in [
    ("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"),
    ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
    ("o", "urn:schemas-microsoft-com:office:office"),
    ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
    ("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
    ("v", "urn:schemas-microsoft-com:vml"),
    ("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"),
    ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
    ("w10", "urn:schemas-microsoft-com:office:word"),
    ("w", W_NS),
    ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
    ("w15", "http://schemas.microsoft.com/office/word/2012/wordml"),
    ("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
    ("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"),
    ("wne", "http://schemas.microsoft.com/office/word/2006/wordml"),
    ("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
]:
    ET.register_namespace(prefix, uri)


@dataclass
class Paper:
    source_path: Path
    title: str
    abstract_paragraphs: list[str]
    keywords: str
    body_blocks: list[dict[str, str]]
    conclusion_paragraphs: list[str]
    references: list[str]


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def parse_markdown_paper(path: Path) -> Paper:
    raw_lines = path.read_text(encoding="utf-8").splitlines()
    lines = [line.strip() for line in raw_lines]
    non_empty = [line for line in lines if line]
    if not non_empty:
        raise ValueError(f"{path} is empty")

    title = non_empty[0].strip("《》")
    try:
        abstract_index = non_empty.index("摘要")
        keywords_index = non_empty.index("关键词")
        references_index = non_empty.index("参考文献")
    except ValueError as exc:
        raise ValueError(f"{path} is missing a required section marker") from exc

    abstract_paragraphs = [line for line in non_empty[abstract_index + 1 : keywords_index] if line]
    keywords_candidates = [line for line in non_empty[keywords_index + 1 : references_index] if line]
    if not keywords_candidates:
        raise ValueError(f"{path} does not contain keywords content")
    keywords = keywords_candidates[0]

    body_source = non_empty[keywords_index + 2 : references_index]
    body_blocks: list[dict[str, str]] = []
    conclusion_paragraphs: list[str] = []
    in_conclusion = False
    for line in body_source:
        if line == "六、结束语":
            in_conclusion = True
            continue
        if in_conclusion:
            if line:
                conclusion_paragraphs.append(line)
            continue
        if H1_RE.match(line):
            body_blocks.append({"type": "h1", "text": line})
        elif H2_RE.match(line):
            body_blocks.append({"type": "h2", "text": line})
        elif line != "关键词":
            body_blocks.append({"type": "p", "text": line})

    references = []
    for line in non_empty[references_index + 1 :]:
        match = REF_RE.match(line)
        if match:
            references.append(f"{match.group(1)} {match.group(2)}")
        else:
            references.append(line)

    return Paper(
        source_path=path,
        title=title,
        abstract_paragraphs=abstract_paragraphs,
        keywords=keywords,
        body_blocks=body_blocks,
        conclusion_paragraphs=conclusion_paragraphs,
        references=references,
    )


def safe_filename(title: str) -> str:
    cleaned = re.sub(r'[<>:"/\\\\|?*]', "_", title)
    return cleaned.strip() or "paper"


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def get_profiles(template_rules: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], dict[str, dict[str, Any]]]:
    paragraph_profiles = {item["profile_id"]: item for item in template_rules["paragraph_profiles"]}
    run_profiles = {item["profile_id"]: item for item in template_rules["run_profiles"]}
    block_rules = {}
    for block in template_rules["agent_format_requirements"]["identified_blocks"]:
        paragraph_profile_id = block.get("paragraph_profile_id")
        if not paragraph_profile_id:
            # Shape-based blocks such as cover text boxes are meaningful for
            # interpretation, but the current generator only formats paragraph
            # content from paragraph profiles.
            continue
        unique_run_ids = []
        for run_id in block.get("run_profile_ids", []):
            if run_id not in unique_run_ids:
                unique_run_ids.append(run_id)
        block_rules[block["block_id"]] = {
            "paragraph_profile_id": paragraph_profile_id,
            "run_profile_ids": unique_run_ids,
        }
    return paragraph_profiles, run_profiles, block_rules


def clear_document_body(document: Any) -> None:
    rng = document.Range(0, document.Content.End - 1)
    rng.Text = ""


def add_paragraph(document: Any) -> Any:
    paragraph = document.Paragraphs.Add(document.Range(document.Content.End - 1, document.Content.End - 1))
    try:
        paragraph.Format.OutlineLevel = WD_OUTLINE_LEVEL_BODY_TEXT
    except Exception:
        pass
    return paragraph


def set_bool_property(target: Any, name: str, value: bool | None) -> None:
    if value is None:
        return
    setattr(target, name, -1 if value else 0)


def apply_run_format(range_obj: Any, run_format: dict[str, Any]) -> None:
    font = range_obj.Font
    fonts = run_format.get("fonts", {})
    ascii_font = FONT_NAME_MAP.get(fonts.get("ascii"), fonts.get("ascii"))
    east_asia_font = FONT_NAME_MAP.get(fonts.get("east_asia"), fonts.get("east_asia"))
    complex_font = FONT_NAME_MAP.get(fonts.get("complex_script"), fonts.get("complex_script"))
    if east_asia_font:
        font.NameFarEast = east_asia_font
    if ascii_font:
        font.Name = ascii_font
        font.NameAscii = ascii_font
        font.NameOther = ascii_font
    if complex_font:
        font.NameBi = complex_font
    if run_format.get("size_pt") is not None:
        font.Size = run_format["size_pt"]
    if run_format.get("size_complex_script_pt") is not None:
        try:
            font.SizeBi = run_format["size_complex_script_pt"]
        except Exception:
            pass
    if "bold" in run_format:
        set_bool_property(font, "Bold", run_format.get("bold"))
    if "italic" in run_format:
        set_bool_property(font, "Italic", run_format.get("italic"))
    underline = run_format.get("underline")
    if underline in UNDERLINE_MAP:
        font.Underline = UNDERLINE_MAP[underline]


def apply_paragraph_format(paragraph: Any, paragraph_profile: dict[str, Any]) -> None:
    fmt = paragraph_profile.get("format", {}).get("paragraph_format", {})
    para_fmt = paragraph.Format
    alignment = fmt.get("alignment")
    if alignment in ALIGNMENT_MAP:
        para_fmt.Alignment = ALIGNMENT_MAP[alignment]
    indent = fmt.get("indent", {})
    if indent.get("left_pt") is not None:
        para_fmt.LeftIndent = indent["left_pt"]
    if indent.get("right_pt") is not None:
        para_fmt.RightIndent = indent["right_pt"]
    if indent.get("first_line_pt") is not None:
        para_fmt.FirstLineIndent = indent["first_line_pt"]
    spacing = fmt.get("spacing", {})
    if spacing.get("before_pt") is not None:
        para_fmt.SpaceBefore = spacing["before_pt"]
    if spacing.get("after_pt") is not None:
        para_fmt.SpaceAfter = spacing["after_pt"]
    if spacing.get("line_pt") is not None:
        para_fmt.LineSpacing = spacing["line_pt"]
    if spacing.get("line_rule") in LINE_SPACING_RULE_MAP:
        para_fmt.LineSpacingRule = LINE_SPACING_RULE_MAP[spacing["line_rule"]]


def fill_paragraph(
    paragraph: Any,
    paragraph_profile: dict[str, Any],
    run_segments: list[tuple[str, dict[str, Any]]],
    *,
    outline_level: int = WD_OUTLINE_LEVEL_BODY_TEXT,
) -> Any:
    apply_paragraph_format(paragraph, paragraph_profile)
    try:
        paragraph.Format.OutlineLevel = outline_level
    except Exception:
        pass
    body_range = paragraph.Range.Duplicate
    body_range.End -= 1
    text = "".join(text for text, _ in run_segments)
    body_range.Text = text

    default_run = paragraph_profile.get("format", {}).get("default_run_format", {})
    body_range = paragraph.Range.Duplicate
    body_range.End -= 1
    apply_run_format(body_range, default_run)

    start = paragraph.Range.Start
    offset = 0
    for segment_text, run_profile in run_segments:
        segment_range = paragraph.Range.Duplicate
        segment_range.Start = start + offset
        segment_range.End = segment_range.Start + len(segment_text)
        apply_run_format(segment_range, run_profile.get("format", {}))
        offset += len(segment_text)
    return paragraph


def append_formatted_paragraph(
    document: Any,
    paragraph_profile: dict[str, Any],
    run_segments: list[tuple[str, dict[str, Any]]],
    *,
    outline_level: int = WD_OUTLINE_LEVEL_BODY_TEXT,
) -> Any:
    paragraph = add_paragraph(document)
    return fill_paragraph(paragraph, paragraph_profile, run_segments, outline_level=outline_level)


def append_paragraph_from_donor(document: Any, donor_paragraph: Any, text: str) -> Any:
    paragraph = add_paragraph(document)
    paragraph.Range.FormattedText = donor_paragraph.Range.FormattedText
    body_range = paragraph.Range.Duplicate
    body_range.End -= 1
    body_range.Text = text
    return paragraph


def append_blank_paragraphs(document: Any, count: int) -> None:
    for _ in range(count):
        add_paragraph(document)


def insert_page_break(document: Any) -> None:
    document.Range(document.Content.End - 1, document.Content.End - 1).InsertBreak(WD_PAGE_BREAK)


def add_bookmark(document: Any, name: str, paragraph: Any) -> None:
    bookmark_range = paragraph.Range.Duplicate
    bookmark_range.End -= 1
    if document.Bookmarks.Exists(name):
        document.Bookmarks(name).Delete()
    document.Bookmarks.Add(name, bookmark_range)


def make_toc_text(label: str, page_number: int) -> str:
    dots = "." * max(20, 58 - len(label))
    return f"{label}{dots}{page_number}"


def build_paper_docx(
    paper: Paper,
    template_doc_path: Path,
    template_rules: dict[str, Any],
    output_path: Path,
) -> None:
    paragraph_profiles, run_profiles, block_rules = get_profiles(template_rules)
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    source_document = None
    donor_document = None
    donor_path = output_path.with_name(f"{output_path.stem}.__donor.docx")
    try:
        if output_path.exists():
            output_path.unlink()
        if donor_path.exists():
            donor_path.unlink()
        source_document = word.Documents.Open(str(template_doc_path.resolve()), ReadOnly=True)
        source_document.SaveAs(str(donor_path.resolve()), FileFormat=16)
        source_document.Close(False)
        source_document = None
        shutil.copyfile(donor_path, output_path)
        donor_document = word.Documents.Open(str(donor_path.resolve()), ReadOnly=True)
        document = word.Documents.Open(str(output_path.resolve()))
        clear_document_body(document)
        donor_indices = {
            block["block_id"]: block["paragraph_index"]
            for block in template_rules["agent_format_requirements"]["identified_blocks"]
            if block.get("paragraph_index") and not block.get("part")
        }

        # Cover
        append_blank_paragraphs(document, 8)
        append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["cover_title"]["paragraph_profile_id"]],
            [
                ("    ", run_profiles["R004"]),
                (f"课题名称：{paper.title}", run_profiles["R005"]),
            ],
        )
        append_blank_paragraphs(document, 4)
        append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["cover_info_line"]["paragraph_profile_id"]],
            [(f"专  业：机电一体化{' ' * 12}", run_profiles["R006"])],
        )
        append_blank_paragraphs(document, 1)
        append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["cover_info_line"]["paragraph_profile_id"]],
            [("姓  名：            ", run_profiles["R006"])],
        )
        append_blank_paragraphs(document, 1)
        append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["cover_info_line"]["paragraph_profile_id"]],
            [("学  号：            ", run_profiles["R006"])],
        )
        append_blank_paragraphs(document, 5)
        append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["cover_school"]["paragraph_profile_id"]],
            [("山东科技职业学院", run_profiles["R007"])],
        )
        append_blank_paragraphs(document, 1)
        append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["cover_date"]["paragraph_profile_id"]],
            [("2026年3月", run_profiles["R007"])],
        )
        insert_page_break(document)

        # TOC
        append_blank_paragraphs(document, 1)
        append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["toc_title"]["paragraph_profile_id"]],
            [("目  录", run_profiles["R008"])],
        )
        append_blank_paragraphs(document, 1)
        toc_entries: list[tuple[Any, str, str]] = []

        def add_toc_placeholder(label: str, bookmark_name: str) -> None:
            paragraph = append_formatted_paragraph(
                document,
                paragraph_profiles[block_rules["toc_entry"]["paragraph_profile_id"]],
                [(make_toc_text(label, 1), run_profiles["R009"])],
            )
            toc_entries.append((paragraph, label, bookmark_name))

        add_toc_placeholder("摘要", "bm_abstract")
        add_toc_placeholder("关键词", "bm_keywords")
        add_toc_placeholder("前言", "bm_preface")
        section_bookmarks: list[tuple[str, str]] = []
        h1_counter = 0
        h2_counter = 0
        for block in paper.body_blocks:
            if block["type"] == "h1":
                h1_counter += 1
                bookmark_name = f"bm_h1_{h1_counter:02d}"
                section_bookmarks.append((bookmark_name, block["text"]))
                add_toc_placeholder(block["text"], bookmark_name)
            elif block["type"] == "h2":
                h2_counter += 1
                bookmark_name = f"bm_h2_{h2_counter:02d}"
                section_bookmarks.append((bookmark_name, block["text"]))
                add_toc_placeholder(block["text"], bookmark_name)
        add_toc_placeholder("结束语", "bm_conclusion")
        add_toc_placeholder("参考文献", "bm_references")
        insert_page_break(document)

        # Abstract and keywords
        append_blank_paragraphs(document, 1)
        abstract_title = append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["abstract_title"]["paragraph_profile_id"]],
            [("摘  要", run_profiles["R011"])],
            outline_level=WD_OUTLINE_LEVEL_1,
        )
        add_bookmark(document, "bm_abstract", abstract_title)
        append_blank_paragraphs(document, 1)
        for index, paragraph_text in enumerate(paper.abstract_paragraphs):
            segments = []
            if index == 0:
                segments.append(("\t", run_profiles["R012"]))
            segments.append((paragraph_text, run_profiles["R013"]))
            append_formatted_paragraph(
                document,
                paragraph_profiles[block_rules["abstract_body"]["paragraph_profile_id"]],
                segments,
            )
        append_blank_paragraphs(document, 1)
        keywords_paragraph = append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["keywords_line"]["paragraph_profile_id"]],
            [
                ("关键词", run_profiles["R014"]),
                (f"：{paper.keywords}", run_profiles["R015"]),
            ],
            outline_level=WD_OUTLINE_LEVEL_BODY_TEXT,
        )
        add_bookmark(document, "bm_keywords", keywords_paragraph)
        append_blank_paragraphs(document, 1)
        preface_title = append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["preface_title"]["paragraph_profile_id"]],
            [("前  言", run_profiles["R016"])],
            outline_level=WD_OUTLINE_LEVEL_1,
        )
        add_bookmark(document, "bm_preface", preface_title)
        append_blank_paragraphs(document, 1)

        # Body
        h1_counter = 0
        h2_counter = 0
        for block in paper.body_blocks:
            if block["type"] == "h1":
                h1_counter += 1
                paragraph = append_formatted_paragraph(
                    document,
                    paragraph_profiles[block_rules["level_1_heading"]["paragraph_profile_id"]],
                    [(block["text"], run_profiles["R017"])],
                    outline_level=WD_OUTLINE_LEVEL_1,
                )
                add_bookmark(document, f"bm_h1_{h1_counter:02d}", paragraph)
            elif block["type"] == "h2":
                h2_counter += 1
                paragraph = append_formatted_paragraph(
                    document,
                    paragraph_profiles[block_rules["level_2_heading"]["paragraph_profile_id"]],
                    [(block["text"], run_profiles["R019"])],
                    outline_level=WD_OUTLINE_LEVEL_2,
                )
                add_bookmark(document, f"bm_h2_{h2_counter:02d}", paragraph)
            else:
                append_formatted_paragraph(
                    document,
                    paragraph_profiles[block_rules["body_text"]["paragraph_profile_id"]],
                    [(block["text"], run_profiles["R018"])],
                    outline_level=WD_OUTLINE_LEVEL_BODY_TEXT,
                )

        append_blank_paragraphs(document, 1)
        conclusion_title = append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["conclusion_title"]["paragraph_profile_id"]],
            [("结束语", run_profiles["R020"])],
            outline_level=WD_OUTLINE_LEVEL_1,
        )
        add_bookmark(document, "bm_conclusion", conclusion_title)
        for paragraph_text in paper.conclusion_paragraphs:
            append_formatted_paragraph(
                document,
                paragraph_profiles[block_rules["body_text"]["paragraph_profile_id"]],
                [(paragraph_text, run_profiles["R018"])],
                outline_level=WD_OUTLINE_LEVEL_BODY_TEXT,
            )

        append_blank_paragraphs(document, 1)
        references_title = append_formatted_paragraph(
            document,
            paragraph_profiles[block_rules["references_title"]["paragraph_profile_id"]],
            [("参考文献", run_profiles["R022"])],
            outline_level=WD_OUTLINE_LEVEL_1,
        )
        add_bookmark(document, "bm_references", references_title)
        append_blank_paragraphs(document, 1)
        for item in paper.references:
            append_formatted_paragraph(
                document,
                paragraph_profiles[block_rules["reference_item"]["paragraph_profile_id"]],
                [(item, run_profiles["R023"])],
                outline_level=WD_OUTLINE_LEVEL_BODY_TEXT,
            )

        document.Repaginate()
        for toc_paragraph, label, bookmark_name in toc_entries:
            if document.Bookmarks.Exists(bookmark_name):
                page_number = int(document.Bookmarks(bookmark_name).Range.Information(WD_INFO_ADJUSTED_PAGE_NUMBER))
            else:
                page_number = 1
            fill_paragraph(
                toc_paragraph,
                paragraph_profiles[block_rules["toc_entry"]["paragraph_profile_id"]],
                [(make_toc_text(label, page_number), run_profiles["R009"])],
                outline_level=WD_OUTLINE_LEVEL_BODY_TEXT,
            )

        document.Save()
    finally:
        if source_document is not None:
            source_document.Close(False)
        if donor_document is not None:
            donor_document.Close(False)
        if document is not None:
            document.Close(False)
        word.Quit()
        if donor_path.exists():
            try:
                donor_path.unlink()
            except OSError:
                pass
        del source_document
        del donor_document
        del document
        del word
        gc.collect()
        time.sleep(1.0)


def round_or_none(value: float | None) -> float | None:
    return None if value is None else round(float(value), 1)


def simplify_paragraph_signature(paragraph_format: dict[str, Any]) -> dict[str, Any]:
    return {
        "alignment": paragraph_format.get("alignment"),
        "first_line_pt": round_or_none(paragraph_format.get("indent", {}).get("first_line_pt")),
        "line_pt": round_or_none(paragraph_format.get("spacing", {}).get("line_pt")),
        "line_rule": paragraph_format.get("spacing", {}).get("line_rule"),
    }


def simplify_run_signature(run_format: dict[str, Any]) -> dict[str, Any]:
    fonts = run_format.get("fonts", {})
    return {
        "ascii": FONT_ALIAS_MAP.get(fonts.get("ascii"), fonts.get("ascii")),
        "east_asia": FONT_ALIAS_MAP.get(fonts.get("east_asia"), fonts.get("east_asia")),
        "size_pt": round_or_none(run_format.get("size_pt")),
        "bold": run_format.get("bold"),
        "underline": run_format.get("underline"),
    }


def font_value_matches(actual: str | None, expected: str | None, block_id: str) -> bool:
    if actual == expected:
        return True
    if expected == "simsun" and actual == "fangsong_gb2312" and block_id in WORD_SIMSUN_SUBSTITUTE_BLOCKS:
        return True
    return False


def run_signature_matches(actual: dict[str, Any], expected: dict[str, Any], block_id: str) -> bool:
    for key, expected_value in expected.items():
        actual_value = actual.get(key)
        if key in {"ascii", "east_asia"}:
            if not font_value_matches(actual_value, expected_value, block_id):
                return False
        elif expected_value != actual_value:
            return False
    return True


def classify_extracted_document(extracted: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    classified: dict[str, list[dict[str, Any]]] = {
        "cover_title": [],
        "cover_info_line": [],
        "cover_school": [],
        "cover_date": [],
        "toc_title": [],
        "toc_entry": [],
        "abstract_title": [],
        "abstract_body": [],
        "keywords_line": [],
        "preface_title": [],
        "level_1_heading": [],
        "level_2_heading": [],
        "body_text": [],
        "conclusion_title": [],
        "references_title": [],
        "reference_item": [],
        "header_primary": [],
        "footer_primary": [],
    }
    main_part = next(part for part in extracted["parts"] if part["part_kind"] == "document")
    non_empty = [paragraph for paragraph in main_part["paragraphs"] if paragraph["text"].strip()]
    mode = "cover"
    for paragraph in non_empty:
        text = paragraph["text"].strip()
        if mode == "toc" and text != "摘  要":
            classified["toc_entry"].append(paragraph)
            continue
        if "课题名称：" in text:
            classified["cover_title"].append(paragraph)
            continue
        if text.startswith("专  业：") or text.startswith("姓  名：") or text.startswith("学  号："):
            classified["cover_info_line"].append(paragraph)
            continue
        if text == "山东科技职业学院":
            classified["cover_school"].append(paragraph)
            continue
        if re.match(r"^\d{4}年\d{1,2}月$", text):
            classified["cover_date"].append(paragraph)
            continue
        if text == "目  录":
            classified["toc_title"].append(paragraph)
            mode = "toc"
            continue
        if text == "摘  要":
            classified["abstract_title"].append(paragraph)
            mode = "abstract"
            continue
        if text.startswith("关键词"):
            classified["keywords_line"].append(paragraph)
            mode = "post_keywords"
            continue
        if text == "前  言":
            classified["preface_title"].append(paragraph)
            mode = "body"
            continue
        if text == "结束语":
            classified["conclusion_title"].append(paragraph)
            mode = "conclusion"
            continue
        if text == "参考文献":
            classified["references_title"].append(paragraph)
            mode = "references"
            continue
        if mode == "toc":
            classified["toc_entry"].append(paragraph)
        elif mode == "abstract":
            classified["abstract_body"].append(paragraph)
        elif mode == "references":
            classified["reference_item"].append(paragraph)
        elif H1_RE.match(text):
            classified["level_1_heading"].append(paragraph)
        elif H2_RE.match(text):
            classified["level_2_heading"].append(paragraph)
        else:
            classified["body_text"].append(paragraph)

    for part in extracted["parts"]:
        if part["part_kind"] == "header":
            for paragraph in part["paragraphs"]:
                if paragraph["text"].strip():
                    classified["header_primary"].append(paragraph)
                    break
        if part["part_kind"] == "footer":
            for paragraph in part["paragraphs"]:
                if paragraph["text"].strip() or any(run.get("field_instruction") for run in paragraph.get("runs", [])):
                    classified["footer_primary"].append(paragraph)
                    break
    return classified


def validate_generated_docx(
    output_path: Path,
    template_rules: dict[str, Any],
) -> dict[str, Any]:
    extracted = extract_document_dict(output_path)
    paragraph_profiles, run_profiles, block_rules = get_profiles(template_rules)
    classified = classify_extracted_document(extracted)
    errors: list[str] = []

    expected_setup = template_rules["page_setup"][0]
    actual_setup = extracted["page_setup"][0]
    for field in ("orientation",):
        if actual_setup.get(field) != expected_setup.get(field):
            errors.append(f"page_setup.{field} expected {expected_setup.get(field)} got {actual_setup.get(field)}")
    for side in ("top_pt", "bottom_pt", "left_pt", "right_pt"):
        if abs(actual_setup["margins"][side] - expected_setup["margins"][side]) > 0.2:
            errors.append(f"page_setup.margins.{side} mismatch")

    required_counts = {
        "cover_title": 1,
        "cover_info_line": 3,
        "cover_school": 1,
        "cover_date": 1,
        "toc_title": 1,
        "toc_entry": 1,
        "abstract_title": 1,
        "abstract_body": 1,
        "keywords_line": 1,
        "preface_title": 1,
        "level_1_heading": 1,
        "level_2_heading": 1,
        "body_text": 1,
        "conclusion_title": 1,
        "references_title": 1,
        "reference_item": 1,
        "header_primary": 1,
        "footer_primary": 1,
    }

    for block_id, minimum in required_counts.items():
        if len(classified.get(block_id, [])) < minimum:
            errors.append(f"{block_id} count is below {minimum}")
            continue
        expected_paragraph = simplify_paragraph_signature(
            paragraph_profiles[block_rules[block_id]["paragraph_profile_id"]]["format"].get("paragraph_format", {})
        )
        expected_run_signatures = [
            simplify_run_signature(run_profiles[run_id]["format"])
            for run_id in block_rules[block_id]["run_profile_ids"]
        ]

        for paragraph in classified[block_id]:
            actual_paragraph = simplify_paragraph_signature(paragraph.get("resolved_paragraph_format", {}))
            for key, value in expected_paragraph.items():
                if value is None:
                    continue
                if actual_paragraph.get(key) != value:
                    errors.append(
                        f"{block_id} paragraph '{paragraph['text'][:30]}' expected {key}={value} got {actual_paragraph.get(key)}"
                    )
            actual_run_signatures = [
                simplify_run_signature(run.get("resolved_format", {}))
                for run in paragraph.get("runs", [])
                if run.get("text") != ""
            ]
            if not actual_run_signatures:
                actual_run_signatures = [simplify_run_signature(paragraph.get("resolved_default_run_format", {}))]
            if not all(any(run_signature_matches(actual, expected, block_id) for expected in expected_run_signatures) for actual in actual_run_signatures):
                errors.append(f"{block_id} run format mismatch in '{paragraph['text'][:30]}'")
            for expected in expected_run_signatures:
                if len(expected_run_signatures) > 1 and not any(run_signature_matches(actual, expected, block_id) for actual in actual_run_signatures):
                    errors.append(f"{block_id} is missing one expected run format in '{paragraph['text'][:30]}'")
                    break

    return {
        "passed": not errors,
        "errors": errors,
        "classified_counts": {key: len(value) for key, value in classified.items()},
    }


def ensure_xml_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(f"w:{tag}", NS)
    if child is None:
        child = ET.SubElement(parent, w(tag))
    return child


def set_xml_bool(rpr: ET.Element, tag: str, value: bool | None) -> None:
    if value is None:
        return
    elem = ensure_xml_child(rpr, tag)
    if value:
        elem.attrib.pop(w("val"), None)
    else:
        elem.set(w("val"), "0")


def apply_run_format_xml(run: ET.Element, run_format: dict[str, Any]) -> None:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(w("rPr"))
        run.insert(0, rpr)
    fonts = run_format.get("fonts", {})
    if fonts:
        rfonts = ensure_xml_child(rpr, "rFonts")
        if fonts.get("ascii"):
            rfonts.set(w("ascii"), fonts["ascii"])
        if fonts.get("east_asia"):
            rfonts.set(w("eastAsia"), fonts["east_asia"])
        if fonts.get("h_ansi"):
            rfonts.set(w("hAnsi"), fonts["h_ansi"])
        if fonts.get("complex_script"):
            rfonts.set(w("cs"), fonts["complex_script"])
        if fonts.get("hint"):
            rfonts.set(w("hint"), fonts["hint"])
    if run_format.get("size_pt") is not None:
        size = ensure_xml_child(rpr, "sz")
        size.set(w("val"), str(int(run_format["size_pt"] * 2)))
    if run_format.get("size_complex_script_pt") is not None:
        size_cs = ensure_xml_child(rpr, "szCs")
        size_cs.set(w("val"), str(int(run_format["size_complex_script_pt"] * 2)))
    set_xml_bool(rpr, "b", run_format.get("bold"))
    set_xml_bool(rpr, "bCs", run_format.get("bold_complex_script"))
    underline = run_format.get("underline")
    if underline is not None:
        u = ensure_xml_child(rpr, "u")
        u.set(w("val"), underline)


def postprocess_generated_docx(source_path: Path, destination_path: Path, template_rules: dict[str, Any]) -> None:
    extracted = extract_document_dict(source_path)
    _, run_profiles, block_rules = get_profiles(template_rules)
    classified = classify_extracted_document(extracted)
    target_blocks = {
        "toc_title",
        "toc_entry",
        "preface_title",
        "body_text",
        "conclusion_title",
    }
    index_to_run_format: dict[int, dict[str, Any]] = {}
    for block_id in target_blocks:
        run_id = block_rules[block_id]["run_profile_ids"][0]
        run_format = run_profiles[run_id]["format"]
        for paragraph in classified.get(block_id, []):
            index_to_run_format[paragraph["index"]] = run_format

    with zipfile.ZipFile(source_path, "r") as archive:
        payload = {name: archive.read(name) for name in archive.namelist()}
    original_document_xml = payload["word/document.xml"].decode("utf-8", errors="replace")

    root = ET.fromstring(payload["word/document.xml"])
    paragraphs = root.findall(".//w:body/w:p", NS)
    for index, run_format in index_to_run_format.items():
        if index - 1 >= len(paragraphs):
            continue
        for run in paragraphs[index - 1].findall("w:r", NS):
            apply_run_format_xml(run, run_format)

    serialized_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")
    original_decl_match = re.match(r"<\?xml[^>]*\?>", original_document_xml)
    original_root_match = re.search(r"<w:document\b[^>]*>", original_document_xml, re.S)
    if original_root_match:
        serialized_xml = re.sub(
            r"<w:document\b[^>]*>",
            original_root_match.group(0),
            serialized_xml,
            count=1,
            flags=re.S,
        )
    if original_decl_match:
        serialized_xml = re.sub(
            r"<\?xml[^>]*\?>",
            original_decl_match.group(0),
            serialized_xml,
            count=1,
        )
    payload["word/document.xml"] = serialized_xml.encode("utf-8")
    if destination_path.exists():
        destination_path.unlink()
    fd, temp_name = tempfile.mkstemp(suffix=".docx", dir=str(destination_path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, content in payload.items():
                archive.writestr(name, content)
        temp_path.replace(destination_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def generate_until_valid(
    paper: Paper,
    template_doc_path: Path,
    template_rules: dict[str, Any],
    output_path: Path,
    attempts: int,
) -> dict[str, Any]:
    last_report = None
    for _ in range(attempts):
        build_paper_docx(paper, template_doc_path, template_rules, output_path)
        last_report = validate_generated_docx(output_path, template_rules)
        if last_report["passed"]:
            return last_report
    raise RuntimeError(f"{output_path.name} failed validation: {last_report['errors']}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate DOCX papers from markdown using template format JSON.")
    parser.add_argument("--template-doc", required=True, help="Path to the source template .doc")
    parser.add_argument("--template-json", required=True, help="Path to the extracted template JSON")
    parser.add_argument("--output-dir", required=True, help="Directory for generated docx files")
    parser.add_argument("markdown_files", nargs="+", help="Markdown paper files")
    args = parser.parse_args()

    template_doc_path = Path(args.template_doc).resolve()
    template_json_path = Path(args.template_json).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    template_rules = load_json(template_json_path)

    summary = []
    for markdown_file in args.markdown_files:
        paper = parse_markdown_paper(Path(markdown_file).resolve())
        output_path = output_dir / f"{safe_filename(paper.title)}.docx"
        report = generate_until_valid(paper, template_doc_path, template_rules, output_path, attempts=3)
        summary.append(
            {
                "source": str(paper.source_path),
                "output": str(output_path),
                "validation": report,
            }
        )

    report_path = output_dir / "validation_report.json"
    report_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
