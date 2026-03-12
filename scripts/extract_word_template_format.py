from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import tempfile
import zipfile
from collections import Counter, defaultdict
from copy import deepcopy
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET

import win32com.client


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

WD_STATISTIC_WORDS = 0
WD_STATISTIC_PAGES = 2
WD_STATISTIC_CHARACTERS = 3
WD_FORMAT_DOCUMENT_DEFAULT = 16

WD_ORIENTATION = {0: "portrait", 1: "landscape"}
HEADER_FOOTER_KIND = {1: "primary", 2: "first_page", 3: "even_pages"}
COM_ALIGNMENT = {
    0: "left",
    1: "center",
    2: "right",
    3: "both",
    4: "distribute",
    5: "justify_medium",
    7: "justify_high",
    8: "justify_low",
}
POINT_SIZE_NAMES = {
    42.0: "初号",
    36.0: "小初",
    26.0: "一号",
    24.0: "小一",
    22.0: "二号",
    18.0: "小二",
    16.0: "三号",
    15.0: "小三",
    14.0: "四号",
    12.0: "小四",
    10.5: "五号",
    9.0: "小五",
}


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def get_attr(elem: ET.Element | None, name: str) -> str | None:
    if elem is None:
        return None
    return elem.attrib.get(w(name))


def pt_to_mm(value: float | int | None) -> float | None:
    if value is None:
        return None
    return round(float(value) * 25.4 / 72.0, 2)


def twips_to_pt(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    return round(int(value) / 20.0, 2)


def half_points_to_pt(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    return round(int(value) / 2.0, 2)


def point_name(value: float | None) -> str | None:
    if value is None:
        return None
    for pt, name in POINT_SIZE_NAMES.items():
        if abs(value - pt) < 0.01:
            return name
    return None


def clean_text(text: str) -> str:
    return text.replace("\r", "").replace("\x07", "")


def com_flag(value: Any) -> bool | None:
    if value is None:
        return None
    try:
        numeric = int(value)
    except (TypeError, ValueError):
        return None
    if numeric == -1:
        return True
    if numeric == 0:
        return False
    return None


def truthy_xml_flag(elem: ET.Element | None) -> bool | None:
    if elem is None:
        return None
    value = get_attr(elem, "val")
    if value is None:
        return True
    return value not in {"0", "false", "off", "False"}


def strip_none(value: Any) -> Any:
    if isinstance(value, dict):
        cleaned = {}
        for key, item in value.items():
            stripped = strip_none(item)
            if stripped is None or stripped == {} or stripped == []:
                continue
            cleaned[key] = stripped
        return cleaned or None
    if isinstance(value, list):
        cleaned = []
        for item in value:
            stripped = strip_none(item)
            if stripped is None or stripped == {} or stripped == []:
                continue
            cleaned.append(stripped)
        return cleaned or None
    return value


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def normalize_signature(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def extract_run_props(rpr: ET.Element | None) -> dict[str, Any]:
    if rpr is None:
        return {}
    fonts = rpr.find("w:rFonts", NS)
    size_pt = half_points_to_pt(get_attr(rpr.find("w:sz", NS), "val"))
    size_cs_pt = half_points_to_pt(get_attr(rpr.find("w:szCs", NS), "val"))
    underline_elem = rpr.find("w:u", NS)
    language = rpr.find("w:lang", NS)
    color = get_attr(rpr.find("w:color", NS), "val")
    props = {
        "fonts": {
            "ascii": get_attr(fonts, "ascii"),
            "east_asia": get_attr(fonts, "eastAsia"),
            "h_ansi": get_attr(fonts, "hAnsi"),
            "complex_script": get_attr(fonts, "cs"),
            "hint": get_attr(fonts, "hint"),
        },
        "size_pt": size_pt,
        "size_name": point_name(size_pt),
        "size_complex_script_pt": size_cs_pt,
        "size_complex_script_name": point_name(size_cs_pt),
        "bold": truthy_xml_flag(rpr.find("w:b", NS)),
        "bold_complex_script": truthy_xml_flag(rpr.find("w:bCs", NS)),
        "italic": truthy_xml_flag(rpr.find("w:i", NS)),
        "italic_complex_script": truthy_xml_flag(rpr.find("w:iCs", NS)),
        "underline": None if underline_elem is None else (get_attr(underline_elem, "val") or "single"),
        "color": None if color in (None, "auto") else color,
        "highlight": get_attr(rpr.find("w:highlight", NS), "val"),
        "strike": truthy_xml_flag(rpr.find("w:strike", NS)),
        "double_strike": truthy_xml_flag(rpr.find("w:dstrike", NS)),
        "emboss": truthy_xml_flag(rpr.find("w:emboss", NS)),
        "imprint": truthy_xml_flag(rpr.find("w:imprint", NS)),
        "outline": truthy_xml_flag(rpr.find("w:outline", NS)),
        "shadow": truthy_xml_flag(rpr.find("w:shadow", NS)),
        "small_caps": truthy_xml_flag(rpr.find("w:smallCaps", NS)),
        "all_caps": truthy_xml_flag(rpr.find("w:caps", NS)),
        "vanish": truthy_xml_flag(rpr.find("w:vanish", NS)),
        "rtl": truthy_xml_flag(rpr.find("w:rtl", NS)),
        "web_hidden": truthy_xml_flag(rpr.find("w:webHidden", NS)),
        "position_half_points": int(get_attr(rpr.find("w:position", NS), "val")) if get_attr(rpr.find("w:position", NS), "val") else None,
        "spacing_twips": int(get_attr(rpr.find("w:spacing", NS), "val")) if get_attr(rpr.find("w:spacing", NS), "val") else None,
        "kern_half_points": int(get_attr(rpr.find("w:kern", NS), "val")) if get_attr(rpr.find("w:kern", NS), "val") else None,
        "vertical_align": get_attr(rpr.find("w:vertAlign", NS), "val"),
        "emphasis_mark": get_attr(rpr.find("w:em", NS), "val"),
        "language": {
            "val": get_attr(language, "val"),
            "east_asia": get_attr(language, "eastAsia"),
            "bidi": get_attr(language, "bidi"),
        },
    }
    return strip_none(props) or {}


def extract_paragraph_props(ppr: ET.Element | None) -> dict[str, Any]:
    if ppr is None:
        return {}
    ind = ppr.find("w:ind", NS)
    spacing = ppr.find("w:spacing", NS)
    num_pr = ppr.find("w:numPr", NS)
    tabs = []
    tabs_parent = ppr.find("w:tabs", NS)
    if tabs_parent is not None:
        for tab in tabs_parent.findall("w:tab", NS):
            tabs.append(
                strip_none(
                    {
                        "alignment": get_attr(tab, "val"),
                        "leader": get_attr(tab, "leader"),
                        "position_twips": int(get_attr(tab, "pos")) if get_attr(tab, "pos") else None,
                        "position_pt": twips_to_pt(get_attr(tab, "pos")),
                    }
                )
            )
    props = {
        "style_id": get_attr(ppr.find("w:pStyle", NS), "val"),
        "alignment": get_attr(ppr.find("w:jc", NS), "val"),
        "indent": {
            "left_twips": int(get_attr(ind, "left")) if get_attr(ind, "left") else None,
            "left_pt": twips_to_pt(get_attr(ind, "left")),
            "right_twips": int(get_attr(ind, "right")) if get_attr(ind, "right") else None,
            "right_pt": twips_to_pt(get_attr(ind, "right")),
            "first_line_twips": int(get_attr(ind, "firstLine")) if get_attr(ind, "firstLine") else None,
            "first_line_pt": twips_to_pt(get_attr(ind, "firstLine")),
            "hanging_twips": int(get_attr(ind, "hanging")) if get_attr(ind, "hanging") else None,
            "hanging_pt": twips_to_pt(get_attr(ind, "hanging")),
            "left_chars": int(get_attr(ind, "leftChars")) if get_attr(ind, "leftChars") else None,
            "right_chars": int(get_attr(ind, "rightChars")) if get_attr(ind, "rightChars") else None,
            "first_line_chars": int(get_attr(ind, "firstLineChars")) if get_attr(ind, "firstLineChars") else None,
            "hanging_chars": int(get_attr(ind, "hangingChars")) if get_attr(ind, "hangingChars") else None,
        },
        "spacing": {
            "before_twips": int(get_attr(spacing, "before")) if get_attr(spacing, "before") else None,
            "before_pt": twips_to_pt(get_attr(spacing, "before")),
            "after_twips": int(get_attr(spacing, "after")) if get_attr(spacing, "after") else None,
            "after_pt": twips_to_pt(get_attr(spacing, "after")),
            "line_twips": int(get_attr(spacing, "line")) if get_attr(spacing, "line") else None,
            "line_pt": twips_to_pt(get_attr(spacing, "line")),
            "line_rule": get_attr(spacing, "lineRule"),
        },
        "outline_level": int(get_attr(ppr.find("w:outlineLvl", NS), "val")) if get_attr(ppr.find("w:outlineLvl", NS), "val") else None,
        "keep_next": truthy_xml_flag(ppr.find("w:keepNext", NS)),
        "keep_lines": truthy_xml_flag(ppr.find("w:keepLines", NS)),
        "page_break_before": truthy_xml_flag(ppr.find("w:pageBreakBefore", NS)),
        "widow_control": truthy_xml_flag(ppr.find("w:widowControl", NS)),
        "contextual_spacing": truthy_xml_flag(ppr.find("w:contextualSpacing", NS)),
        "suppress_auto_hyphens": truthy_xml_flag(ppr.find("w:suppressAutoHyphens", NS)),
        "snap_to_grid": truthy_xml_flag(ppr.find("w:snapToGrid", NS)),
        "tabs": tabs,
        "numbering": {
            "num_id": get_attr(num_pr.find("w:numId", NS), "val") if num_pr is not None else None,
            "level": int(get_attr(num_pr.find("w:ilvl", NS), "val")) if num_pr is not None and get_attr(num_pr.find("w:ilvl", NS), "val") else None,
        },
    }
    return strip_none(props) or {}


def iter_runs(node: ET.Element) -> Iterable[ET.Element]:
    for child in list(node):
        if child.tag == w("r"):
            yield child
        elif child.tag in {w("hyperlink"), w("smartTag"), w("sdt"), w("fldSimple"), w("customXml")}:
            yield from iter_runs(child)


def extract_run_content(run: ET.Element) -> dict[str, Any]:
    text_parts = []
    field_instruction_parts = []
    control_tokens = []
    field_char_type = None
    for child in list(run):
        if child.tag == w("t"):
            text_parts.append(child.text or "")
        elif child.tag == w("tab"):
            text_parts.append("\t")
            control_tokens.append("tab")
        elif child.tag in {w("br"), w("cr")}:
            text_parts.append("\n")
            control_tokens.append("line_break")
        elif child.tag == w("instrText"):
            field_instruction_parts.append(child.text or "")
        elif child.tag == w("fldChar"):
            field_char_type = get_attr(child, "fldCharType")
        elif child.tag == w("noBreakHyphen"):
            text_parts.append("-")
        elif child.tag == w("softHyphen"):
            text_parts.append("")
        elif child.tag == w("sym"):
            text_parts.append(f"[symbol:{get_attr(child, 'font')}:{get_attr(child, 'char')}]")
            control_tokens.append("symbol")
    payload = {
        "text": "".join(text_parts),
        "field_instruction": "".join(field_instruction_parts) or None,
        "field_char_type": field_char_type,
        "controls": control_tokens,
    }
    return strip_none(payload) or {"text": ""}


def load_styles(zf: zipfile.ZipFile) -> dict[str, Any]:
    payload = {
        "doc_defaults": {"paragraph": {}, "run": {}},
        "styles": {},
    }
    if "word/styles.xml" not in zf.namelist():
        return payload
    root = ET.fromstring(zf.read("word/styles.xml"))
    defaults = root.find("w:docDefaults", NS)
    if defaults is not None:
        payload["doc_defaults"]["paragraph"] = extract_paragraph_props(defaults.find("w:pPrDefault/w:pPr", NS))
        payload["doc_defaults"]["run"] = extract_run_props(defaults.find("w:rPrDefault/w:rPr", NS))
    for style in root.findall("w:style", NS):
        style_id = get_attr(style, "styleId")
        payload["styles"][style_id] = strip_none(
            {
                "style_id": style_id,
                "name": get_attr(style.find("w:name", NS), "val"),
                "type": get_attr(style, "type"),
                "based_on": get_attr(style.find("w:basedOn", NS), "val"),
                "paragraph_format": extract_paragraph_props(style.find("w:pPr", NS)),
                "run_format": extract_run_props(style.find("w:rPr", NS)),
            }
        ) or {"style_id": style_id}
    return payload


def resolve_style(style_id: str | None, style_catalog: dict[str, Any], cache: dict[str, Any]) -> dict[str, Any]:
    if not style_id:
        return {"paragraph_format": {}, "run_format": {}}
    if style_id in cache:
        return cache[style_id]
    current = style_catalog["styles"].get(style_id, {})
    parent_id = current.get("based_on")
    parent = resolve_style(parent_id, style_catalog, cache) if parent_id else {"paragraph_format": {}, "run_format": {}}
    resolved = {
        "paragraph_format": deep_merge(parent.get("paragraph_format", {}), current.get("paragraph_format", {})),
        "run_format": deep_merge(parent.get("run_format", {}), current.get("run_format", {})),
    }
    cache[style_id] = resolved
    return resolved


def parse_part(
    part_name: str,
    xml_bytes: bytes,
    part_kind: str,
    style_catalog: dict[str, Any],
    style_cache: dict[str, Any],
    paragraph_profiles: dict[str, dict[str, Any]],
    run_profiles: dict[str, dict[str, Any]],
    paragraph_examples: defaultdict[str, list[str]],
    run_examples: defaultdict[str, list[str]],
) -> list[dict[str, Any]]:
    root = ET.fromstring(xml_bytes)
    body = root.find("w:body", NS)
    container = body if body is not None else root
    paragraphs = []
    for index, paragraph in enumerate(container.findall(".//w:p", NS), start=1):
        ppr = paragraph.find("w:pPr", NS)
        direct_paragraph = extract_paragraph_props(ppr)
        style_id = direct_paragraph.pop("style_id", None)
        paragraph_run_defaults = extract_run_props(ppr.find("w:rPr", NS) if ppr is not None else None)
        style_resolved = resolve_style(style_id, style_catalog, style_cache)
        resolved_paragraph = deep_merge(style_catalog["doc_defaults"]["paragraph"], style_resolved["paragraph_format"])
        resolved_paragraph = deep_merge(resolved_paragraph, direct_paragraph)
        resolved_default_run = deep_merge(style_catalog["doc_defaults"]["run"], style_resolved["run_format"])
        resolved_default_run = deep_merge(resolved_default_run, paragraph_run_defaults)

        paragraph_profile_payload = strip_none(
            {
                "paragraph_format": resolved_paragraph,
                "default_run_format": resolved_default_run,
            }
        ) or {"paragraph_format": {}, "default_run_format": {}}
        paragraph_signature = normalize_signature(paragraph_profile_payload)
        if paragraph_signature not in paragraph_profiles:
            paragraph_profiles[paragraph_signature] = {
                "profile_id": f"P{len(paragraph_profiles) + 1:03d}",
                "usage_count": 0,
                "format": paragraph_profile_payload,
            }
        paragraph_profiles[paragraph_signature]["usage_count"] += 1

        runs = []
        paragraph_text_parts = []
        for run_index, run in enumerate(iter_runs(paragraph), start=1):
            direct_run = extract_run_props(run.find("w:rPr", NS))
            resolved_run = deep_merge(resolved_default_run, direct_run)
            run_signature = normalize_signature(resolved_run)
            if run_signature not in run_profiles:
                run_profiles[run_signature] = {
                    "profile_id": f"R{len(run_profiles) + 1:03d}",
                    "usage_count": 0,
                    "format": resolved_run,
                }
            run_profiles[run_signature]["usage_count"] += 1
            content = extract_run_content(run)
            text = content.get("text", "")
            paragraph_text_parts.append(text)
            runs.append(
                strip_none(
                    {
                        "index": run_index,
                        "text": text,
                        "profile_id": run_profiles[run_signature]["profile_id"],
                        "resolved_format": resolved_run,
                        "field_instruction": content.get("field_instruction"),
                        "field_char_type": content.get("field_char_type"),
                        "controls": content.get("controls"),
                    }
                ) or {"index": run_index, "text": text, "profile_id": run_profiles[run_signature]["profile_id"]}
            )
            if text and len(run_examples[run_signature]) < 3:
                run_examples[run_signature].append(text[:60])

        paragraph_text = "".join(paragraph_text_parts)
        paragraphs.append(
            strip_none(
                {
                    "part": part_name,
                    "part_kind": part_kind,
                    "index": index,
                    "text": paragraph_text,
                    "is_empty": paragraph_text.strip() == "",
                    "style_id": style_id,
                    "style_name": style_catalog["styles"].get(style_id, {}).get("name") if style_id else None,
                    "paragraph_profile_id": paragraph_profiles[paragraph_signature]["profile_id"],
                    "resolved_paragraph_format": resolved_paragraph,
                    "resolved_default_run_format": resolved_default_run,
                    "runs": runs,
                }
            ) or {
                "part": part_name,
                "part_kind": part_kind,
                "index": index,
                "text": paragraph_text,
                "paragraph_profile_id": paragraph_profiles[paragraph_signature]["profile_id"],
                "runs": runs,
            }
        )
        if paragraph_text.strip() and len(paragraph_examples[paragraph_signature]) < 3:
            paragraph_examples[paragraph_signature].append(paragraph_text[:80])
    return paragraphs


def summarize_profiles(
    profiles: dict[str, dict[str, Any]],
    examples: defaultdict[str, list[str]],
) -> list[dict[str, Any]]:
    reverse_lookup = {value["profile_id"]: key for key, value in profiles.items()}
    ordered = sorted(profiles.values(), key=lambda item: (-item["usage_count"], item["profile_id"]))
    summary = []
    for item in ordered:
        signature = reverse_lookup[item["profile_id"]]
        summary.append(
            {
                "profile_id": item["profile_id"],
                "usage_count": item["usage_count"],
                "examples": examples.get(signature, []),
                "format": item["format"],
            }
        )
    return summary


def find_first_content_run(paragraph: dict[str, Any]) -> dict[str, Any] | None:
    for run in paragraph.get("runs", []):
        if (run.get("text") or "").strip() or run.get("field_instruction"):
            return run
    return None


def build_paragraph_block(
    block_id: str,
    name: str,
    paragraph: dict[str, Any] | None,
    *,
    classification: str,
    notes: list[str] | None = None,
    format_requirements: dict[str, Any] | None = None,
    related_paragraphs: list[dict[str, Any]] | None = None,
) -> dict[str, Any] | None:
    if paragraph is None:
        return None
    primary_run = find_first_content_run(paragraph)
    return strip_none(
        {
            "block_id": block_id,
            "name": name,
            "classification": classification,
            "source_kind": "paragraph",
            "paragraph_index": paragraph["index"],
            "paragraph_profile_id": paragraph["paragraph_profile_id"],
            "run_profile_ids": [run["profile_id"] for run in paragraph.get("runs", [])],
            "example_text": paragraph["text"][:120],
            "related_paragraph_indices": [item["index"] for item in related_paragraphs or [] if item["index"] != paragraph["index"]],
            "related_example_texts": [item["text"][:120] for item in related_paragraphs or [] if item["index"] != paragraph["index"]],
            "resolved_format_snapshot": {
                "paragraph_format": paragraph.get("resolved_paragraph_format"),
                "primary_run_format": primary_run.get("resolved_format") if primary_run else None,
            },
            "format_requirements": format_requirements,
            "notes": notes,
        }
    )


def build_shape_block(
    block_id: str,
    name: str,
    shape: dict[str, Any] | None,
    *,
    classification: str,
    notes: list[str] | None = None,
    format_requirements: dict[str, Any] | None = None,
) -> dict[str, Any] | None:
    if shape is None:
        return None
    return strip_none(
        {
            "block_id": block_id,
            "name": name,
            "classification": classification,
            "source_kind": "shape_textbox",
            "shape_index": shape.get("index"),
            "shape_name": shape.get("name"),
            "example_text": (shape.get("text") or "")[:120],
            "shape_box": shape.get("box"),
            "resolved_format_snapshot": {
                "paragraph_format": shape.get("paragraph_format"),
                "primary_run_format": shape.get("run_format"),
                "appearance": shape.get("appearance"),
            },
            "format_requirements": format_requirements,
            "notes": notes,
        }
    )


def analyze_document_structure(
    main_paragraphs: list[dict[str, Any]],
    all_parts: list[dict[str, Any]],
    word_snapshot: dict[str, Any],
) -> dict[str, Any]:
    def find_paragraph(predicate) -> dict[str, Any] | None:
        for paragraph in main_paragraphs:
            if predicate(paragraph):
                return paragraph
        return None

    cover_info_lines = [paragraph for paragraph in main_paragraphs if paragraph["text"].startswith(("专  业：", "姓  名：", "学  号："))]
    shape_title = next(
        (
            shape
            for shape in word_snapshot.get("shape_textboxes", [])
            if "毕业大作业" in re.sub(r"\s+", "", shape.get("text") or "")
        ),
        None,
    )
    blocks = [
        build_shape_block(
            "cover_main_title_textbox",
            "封面主标题文本框",
            shape_title,
            classification="cover",
            format_requirements={
                "placement": "文本框",
                "alignment": "center",
                "font_family": "黑体",
                "font_size_pt": 26.0,
                "font_size_name": "一号",
                "bold": True,
                "border": False,
                "fill": False,
            },
            notes=[
                "原始 .doc 中该标题位于独立文本框内，不是普通段落。",
                "生成 Word 时不要把它并入正文标题样式。",
            ],
        ),
        build_paragraph_block(
            "cover_title",
            "封面课题名称行",
            find_paragraph(lambda p: "课题名称" in p["text"]),
            classification="cover",
            format_requirements={
                "alignment": "both",
                "font_family": "黑体",
                "font_size_pt": 24.0,
                "font_size_name": "小一",
                "bold": True,
                "underline": "single",
                "placeholder_style": "通过下划线留空填写",
            },
            notes=[
                "这是一条封面填写项，不是论文标题级别。",
                "行首带手工空格，实际视觉位置不能仅靠样式名推断。",
            ],
        ),
        build_paragraph_block(
            "cover_info_line",
            "封面信息行",
            find_paragraph(lambda p: p["text"].startswith("专  业：")),
            classification="cover",
            related_paragraphs=cover_info_lines,
            format_requirements={
                "alignment": "both",
                "font_family": "黑体",
                "font_size_pt": 18.0,
                "font_size_name": "小二",
                "bold": True,
                "underline": "single",
                "first_line_indent_pt": 72.3,
                "same_format_applies_to": ["姓  名：", "学  号："],
            },
            notes=[
                "专业、姓名、学号三行是同一组封面元信息填写项。",
                "位置依赖首行缩进和手工留白，不属于正文层级标题。",
            ],
        ),
        build_paragraph_block(
            "cover_school",
            "封面学校名称",
            find_paragraph(lambda p: p["text"] == "山东科技职业学院"),
            classification="cover",
            format_requirements={
                "alignment": "center",
                "font_family": "黑体",
                "font_size_pt": 16.0,
                "font_size_name": "三号",
                "bold": True,
                "line_spacing_pt": 20.0,
                "line_spacing_rule": "exact",
            },
            notes=["位于封面底部，使用居中段落，不是页眉内容。"],
        ),
        build_paragraph_block(
            "cover_date",
            "封面日期",
            find_paragraph(lambda p: "年" in p["text"] and "月" in p["text"] and p["index"] < 32),
            classification="cover",
            format_requirements={
                "alignment": "center",
                "font_family": "黑体",
                "font_size_pt": 16.0,
                "font_size_name": "三号",
                "bold": True,
                "line_spacing_pt": 20.0,
                "line_spacing_rule": "exact",
            },
            notes=["日期行为封面底部占位，不应和正文日期字段混淆。"],
        ),
        build_paragraph_block(
            "toc_title",
            "目录标题",
            find_paragraph(lambda p: p["text"] == "目  录"),
            classification="toc",
            format_requirements={
                "alignment": "center",
                "font_family": "宋体",
                "font_size_pt": 15.0,
                "font_size_name": "小三",
                "bold": True,
                "line_spacing_pt": 20.0,
                "line_spacing_rule": "exact",
            },
            notes=["目录标题是普通段落，不是自动目录域标题。"],
        ),
        build_paragraph_block(
            "toc_entry",
            "目录条目",
            find_paragraph(lambda p: "................................" in p["text"] and p["index"] < 63),
            classification="toc",
            format_requirements={
                "alignment": "both",
                "font_family": "宋体",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "bold": False,
                "line_spacing_pt": 23.0,
                "line_spacing_rule": "exact",
                "toc_generation": "手工目录条目与手工点线，不是 TOC 域",
            },
            notes=[
                "目录页由手工文本组成，点线和页码都是普通字符。",
                "不要改成 Word 自动目录，否则格式和页码表现都会变化。",
            ],
        ),
        build_paragraph_block(
            "abstract_title",
            "摘要标题",
            find_paragraph(lambda p: p["text"] == "摘  要"),
            classification="front_matter",
            format_requirements={
                "alignment": "center",
                "font_family": "仿宋_GB2312",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "bold": True,
                "line_spacing_pt": 20.0,
                "line_spacing_rule": "exact",
            },
            notes=["摘要标题使用仿宋加粗居中，与正文一级标题体系不同。"],
        ),
        build_paragraph_block(
            "abstract_body",
            "摘要正文",
            find_paragraph(lambda p: p["text"].startswith("\t在机械工程领域")),
            classification="front_matter",
            format_requirements={
                "alignment": "both",
                "font_family": "仿宋_GB2312",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "line_spacing_pt": 18.0,
                "line_spacing_rule": "auto",
                "indent_behavior": "段首使用 Tab 而不是标准首行缩进",
            },
            notes=["摘要正文开头是制表符缩进，生成时如果改成普通首行缩进，视觉会不同。"],
        ),
        build_paragraph_block(
            "keywords_line",
            "关键词行",
            find_paragraph(lambda p: p["text"].startswith("关键词：")),
            classification="front_matter",
            format_requirements={
                "alignment": "both",
                "font_family": "仿宋",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "line_spacing_pt": 18.0,
                "line_spacing_rule": "auto",
                "mixed_bold": "仅“关键词”三字加粗，后续关键词常规",
            },
            notes=["这是段内局部加粗格式，不能整段统一套粗体。"],
        ),
        build_paragraph_block(
            "preface_title",
            "前言标题",
            find_paragraph(lambda p: p["text"] == "前  言"),
            classification="front_matter",
            format_requirements={
                "alignment": "center",
                "font_family": "宋体",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "bold": True,
                "line_spacing_pt": 28.0,
                "line_spacing_rule": "exact",
            },
            notes=["正文开始前的过渡标题，仍然不是内置标题样式。"],
        ),
        build_paragraph_block(
            "level_1_heading",
            "正文一级标题",
            find_paragraph(lambda p: re.match(r"^[一二三四五六七八九十]+、", p["text"] or "") is not None and p["index"] > 70),
            classification="body",
            format_requirements={
                "alignment": "both",
                "font_family": "黑体",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "bold": False,
                "line_spacing_pt": 28.0,
                "line_spacing_rule": "exact",
                "first_line_indent_pt": 28.0,
            },
            notes=[
                "一级标题依赖直接格式，不要套用 Word 的 heading 1。",
                "该模板的一、二级标题都不是通过段落样式区分的。",
            ],
        ),
        build_paragraph_block(
            "level_2_heading",
            "正文二级标题",
            find_paragraph(lambda p: p["text"].startswith("（一）") and p["index"] > 70),
            classification="body",
            format_requirements={
                "alignment": "both",
                "font_family": "楷体",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "bold": True,
                "line_spacing_pt": 28.0,
                "line_spacing_rule": "exact",
                "first_line_indent_pt": 28.1,
            },
            notes=["二级标题是楷体加粗，不应误判成一级标题或正文。"],
        ),
        build_paragraph_block(
            "body_text",
            "正文段落",
            find_paragraph(lambda p: p["text"].startswith("机电一体化是指") or p["text"].startswith("    综上所述")),
            classification="body",
            format_requirements={
                "alignment": "both",
                "font_family": "宋体",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "bold": False,
                "line_spacing_pt": 28.0,
                "line_spacing_rule": "exact",
                "first_line_indent_pt": 28.0,
                "heading_detection_rule": "数字小点项如 1.数字化 仍按正文格式处理",
            },
            notes=["不要把正文段落或 1.数字化 这种小点项自动升级成标题样式。"],
        ),
        build_paragraph_block(
            "conclusion_title",
            "结束语标题",
            find_paragraph(lambda p: p["text"] == "结束语"),
            classification="body",
            format_requirements={
                "alignment": "center",
                "font_family": "宋体",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "bold": True,
                "line_spacing_pt": 28.0,
                "line_spacing_rule": "exact",
            },
            notes=["结束语标题与前言标题风格接近，但属于正文末尾标题。"],
        ),
        build_paragraph_block(
            "references_title",
            "参考文献标题",
            find_paragraph(lambda p: p["text"] == "参考文献"),
            classification="references",
            format_requirements={
                "alignment": "center",
                "font_family": "仿宋_GB2312",
                "font_size_pt": 16.0,
                "font_size_name": "三号",
                "bold": True,
                "line_spacing_pt": 18.0,
                "line_spacing_rule": "auto",
            },
            notes=["参考文献标题单独使用三号仿宋加粗，和前文标题体系不同。"],
        ),
        build_paragraph_block(
            "reference_item",
            "参考文献条目",
            find_paragraph(lambda p: re.match(r"^\d+\s", p["text"] or "") is not None and p["index"] > 93),
            classification="references",
            format_requirements={
                "alignment": "both",
                "font_family": "仿宋_GB2312",
                "font_size_pt": 14.0,
                "font_size_name": "四号",
                "bold": False,
                "line_spacing_pt": 18.0,
                "line_spacing_rule": "auto",
                "first_line_indent_pt": 28.0,
                "list_behavior": "手工编号，不是自动编号列表",
            },
            notes=["参考文献编号和内容是普通文本，生成时不要切成 Word 自动列表。"],
        ),
    ]

    header_block = None
    footer_block = None
    for part in all_parts:
        if part["part_kind"] == "header":
            paragraph = next((item for item in part["paragraphs"] if item["text"].strip()), None)
            if paragraph:
                header_block = build_paragraph_block(
                    "header_primary",
                    "页眉",
                    paragraph,
                    classification="header_footer",
                    format_requirements={
                        "alignment": "both",
                        "font_family_east_asia": "宋体",
                        "font_family_ascii": "Times New Roman",
                        "font_size_pt": 10.5,
                        "font_size_name": "五号",
                        "scope": "全篇统一页眉",
                    },
                    notes=["文档只有 1 个节，且未启用首页不同页眉，因此封面/目录/正文共用同一页眉。"],
                )
                header_block["part"] = part["part"]
                break
    for part in all_parts:
        if part["part_kind"] == "footer":
            paragraph = next((item for item in part["paragraphs"] if item["text"].strip() or any(run.get("field_instruction") for run in item.get("runs", []))), None)
            if paragraph:
                footer_field = next((run.get("field_instruction") for run in paragraph.get("runs", []) if run.get("field_instruction")), None)
                footer_block = build_paragraph_block(
                    "footer_primary",
                    "页脚页码",
                    paragraph,
                    classification="header_footer",
                    format_requirements={
                        "alignment": "left",
                        "font_family_east_asia": "宋体",
                        "font_family_ascii": "Times New Roman",
                        "font_size_pt": 9.0,
                        "font_size_name": "小五",
                        "field": "PAGE",
                        "scope": "全篇统一页脚页码域",
                    },
                    notes=["页脚使用 PAGE 域，不是普通静态数字；同样没有首页不同设置。"],
                )
                footer_block["part"] = part["part"]
                if footer_block is not None:
                    footer_block["example_text"] = "PAGE 域（页码）" if footer_field and "PAGE" in footer_field.upper() else paragraph["text"][:120]
                break

    identified_blocks = [item for item in blocks if item is not None]
    if header_block:
        identified_blocks.append(header_block)
    if footer_block:
        identified_blocks.append(footer_block)

    return {
        "recognition_notes": [
            "该模板绝大多数内容共用段落样式“正文”，不能只根据 style_name 判断标题、正文或目录。",
            "封面主标题“毕业大作业”位于文本框中，不属于普通 document.xml 段落。",
            "目录是手工输入的目录条目与点线，不是 Word 自动 TOC 域。",
            "封面、摘要和正文中存在大量空段落、手工空格、Tab、下划线占位，版式推断必须结合直接格式和文本位置。",
        ],
        "generation_constraints": [
            "不要把封面填写项、目录条目、关键词行统一映射为标题样式。",
            "不要自动把目录替换成 TOC 域，除非明确要放弃模板原貌。",
            "一级标题、二级标题、正文段落主要依赖字体和行距区分，而不是 Word heading 样式。",
            "参考文献条目和目录条目都应保留为普通文本，不要转换为自动编号或自动目录。",
        ],
        "layout_notes": {
            "section_count": len(word_snapshot.get("page_setup", [])),
            "cover_uses_same_header_footer_as_body": bool(word_snapshot.get("page_setup")) and not any(
                section.get("different_first_page_header_footer") for section in word_snapshot.get("page_setup", [])
            ),
            "different_first_page_header_footer": any(
                section.get("different_first_page_header_footer") for section in word_snapshot.get("page_setup", [])
            ),
            "odd_and_even_pages_header_footer": any(
                section.get("odd_and_even_pages_header_footer") for section in word_snapshot.get("page_setup", [])
            ),
            "toc_kind": "manual",
        },
        "recommended_sequence": [
            "cover_main_title_textbox",
            "cover_title",
            "cover_info_line",
            "cover_school",
            "cover_date",
            "toc_title",
            "toc_entry",
            "abstract_title",
            "abstract_body",
            "keywords_line",
            "preface_title",
            "level_1_heading",
            "level_2_heading",
            "body_text",
            "conclusion_title",
            "references_title",
            "reference_item",
            "header_primary",
            "footer_primary",
        ],
        "identified_blocks": identified_blocks,
    }


def build_format_summary(main_paragraphs: list[dict[str, Any]]) -> dict[str, Any]:
    font_counter = Counter()
    size_counter = Counter()
    alignment_counter = Counter()
    line_spacing_counter = Counter()
    for paragraph in main_paragraphs:
        alignment = paragraph.get("resolved_paragraph_format", {}).get("alignment")
        if alignment:
            alignment_counter[alignment] += 1
        line_pt = paragraph.get("resolved_paragraph_format", {}).get("spacing", {}).get("line_pt")
        if line_pt:
            line_spacing_counter[str(line_pt)] += 1
        for run in paragraph.get("runs", []):
            fonts = run.get("resolved_format", {}).get("fonts", {})
            font_name = fonts.get("east_asia") or fonts.get("ascii")
            if font_name:
                font_counter[font_name] += 1
            size_pt = run.get("resolved_format", {}).get("size_pt")
            if size_pt:
                size_counter[str(size_pt)] += 1
    return {
        "dominant_fonts": [{"name": name, "count": count} for name, count in font_counter.most_common()],
        "dominant_font_sizes_pt": [{"size_pt": float(size), "size_name": point_name(float(size)), "count": count} for size, count in size_counter.most_common()],
        "paragraph_alignments": [{"alignment": alignment, "count": count} for alignment, count in alignment_counter.most_common()],
        "line_spacing_pt": [{"line_pt": float(value), "count": count} for value, count in line_spacing_counter.most_common()],
        "summary_notes": [
            "dominant_* 统计仅反映字符和显式段落格式的分布，不能直接代表语义层级。",
            "本模板语义判断应优先结合文本内容、直接格式、文本框与页内位置。",
        ],
    }


def extract_from_word(source_path: Path, temp_docx_path: Path) -> dict[str, Any]:
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        document = word.Documents.Open(str(source_path), ReadOnly=True)
        stats = {
            "pages": int(document.ComputeStatistics(WD_STATISTIC_PAGES)),
            "words": int(document.ComputeStatistics(WD_STATISTIC_WORDS)),
            "characters": int(document.ComputeStatistics(WD_STATISTIC_CHARACTERS)),
            "paragraphs": int(document.Paragraphs.Count),
            "sections": int(document.Sections.Count),
        }
        sections = []
        for index in range(1, document.Sections.Count + 1):
            section = document.Sections(index)
            page_setup = section.PageSetup
            sections.append(
                {
                    "index": index,
                    "orientation": WD_ORIENTATION.get(int(page_setup.Orientation), str(page_setup.Orientation)),
                    "page_size": {
                        "width_pt": round(float(page_setup.PageWidth), 2),
                        "height_pt": round(float(page_setup.PageHeight), 2),
                        "width_mm": pt_to_mm(page_setup.PageWidth),
                        "height_mm": pt_to_mm(page_setup.PageHeight),
                    },
                    "margins": {
                        "top_pt": round(float(page_setup.TopMargin), 2),
                        "bottom_pt": round(float(page_setup.BottomMargin), 2),
                        "left_pt": round(float(page_setup.LeftMargin), 2),
                        "right_pt": round(float(page_setup.RightMargin), 2),
                        "top_mm": pt_to_mm(page_setup.TopMargin),
                        "bottom_mm": pt_to_mm(page_setup.BottomMargin),
                        "left_mm": pt_to_mm(page_setup.LeftMargin),
                        "right_mm": pt_to_mm(page_setup.RightMargin),
                    },
                    "header_distance_pt": round(float(page_setup.HeaderDistance), 2),
                    "footer_distance_pt": round(float(page_setup.FooterDistance), 2),
                    "header_distance_mm": pt_to_mm(page_setup.HeaderDistance),
                    "footer_distance_mm": pt_to_mm(page_setup.FooterDistance),
                    "different_first_page_header_footer": bool(page_setup.DifferentFirstPageHeaderFooter),
                    "odd_and_even_pages_header_footer": bool(page_setup.OddAndEvenPagesHeaderFooter),
                }
            )
        header_footer = []
        for section_index in range(1, document.Sections.Count + 1):
            section = document.Sections(section_index)
            for collection_name, collection in (("headers", section.Headers), ("footers", section.Footers)):
                for item_index in range(1, collection.Count + 1):
                    item = collection(item_index)
                    header_footer.append(
                        {
                            "section": section_index,
                            "collection": collection_name,
                            "kind": HEADER_FOOTER_KIND.get(item_index, str(item_index)),
                            "exists": bool(clean_text(item.Range.Text).strip()),
                            "text": clean_text(item.Range.Text),
                        }
                    )
        shape_textboxes = []
        for shape_index in range(1, document.Shapes.Count + 1):
            shape = document.Shapes(shape_index)
            text = ""
            try:
                if shape.TextFrame.HasText:
                    text = clean_text(shape.TextFrame.TextRange.Text)
            except Exception:
                text = ""
            if not text.strip():
                continue
            paragraph_format = None
            run_format = None
            try:
                text_range = shape.TextFrame.TextRange
                paragraph_format = strip_none(
                    {
                        "alignment": COM_ALIGNMENT.get(int(text_range.ParagraphFormat.Alignment), str(text_range.ParagraphFormat.Alignment)),
                        "line_spacing_pt": round(float(text_range.ParagraphFormat.LineSpacing), 2),
                    }
                )
                run_format = strip_none(
                    {
                        "font_family_east_asia": text_range.Font.NameFarEast,
                        "font_family_ascii": text_range.Font.NameAscii,
                        "font_size_pt": round(float(text_range.Font.Size), 2),
                        "font_size_name": point_name(round(float(text_range.Font.Size), 2)),
                        "bold": com_flag(text_range.Font.Bold),
                        "italic": com_flag(text_range.Font.Italic),
                    }
                )
            except Exception:
                paragraph_format = None
                run_format = None
            shape_textboxes.append(
                strip_none(
                    {
                        "index": shape_index,
                        "name": shape.Name,
                        "type": int(shape.Type),
                        "text": text,
                        "box": {
                            "left_pt": round(float(shape.Left), 2),
                            "top_pt": round(float(shape.Top), 2),
                            "width_pt": round(float(shape.Width), 2),
                            "height_pt": round(float(shape.Height), 2),
                        },
                        "paragraph_format": paragraph_format,
                        "run_format": run_format,
                        "appearance": {
                            "border": bool(shape.Line.Visible),
                            "fill": bool(shape.Fill.Visible),
                        },
                    }
                )
            )
        if temp_docx_path.exists():
            temp_docx_path.unlink()
        document.SaveAs(str(temp_docx_path), FileFormat=WD_FORMAT_DOCUMENT_DEFAULT)
        return {
            "document_stats": stats,
            "page_setup": sections,
            "header_footer_text_snapshot": header_footer,
            "shape_textboxes": shape_textboxes,
        }
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract Word template formatting into JSON.")
    parser.add_argument("source", help="Path to the source .doc/.docx file")
    parser.add_argument("output", help="Path to write the JSON output")
    args = parser.parse_args()

    source_path = Path(args.source).resolve()
    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result = extract_document_dict(source_path)
    with output_path.open("w", encoding="utf-8") as file:
        json.dump(result, file, ensure_ascii=False, indent=2)


def extract_document_dict(source_path: Path) -> dict[str, Any]:
    source_path = Path(source_path).resolve()
    temp_docx_path = Path(tempfile.gettempdir()) / "codex_template_format_extract.docx"
    word_snapshot = extract_from_word(source_path, temp_docx_path)
    paragraph_profiles: dict[str, dict[str, Any]] = {}
    run_profiles: dict[str, dict[str, Any]] = {}
    paragraph_examples: defaultdict[str, list[str]] = defaultdict(list)
    run_examples: defaultdict[str, list[str]] = defaultdict(list)
    style_cache: dict[str, Any] = {}

    try:
        with zipfile.ZipFile(temp_docx_path) as zf:
            style_catalog = load_styles(zf)
            parts_to_parse = [("word/document.xml", "document", zf.read("word/document.xml"))]
            for name in sorted(item for item in zf.namelist() if re.fullmatch(r"word/header\d+\.xml", item)):
                parts_to_parse.append((name, "header", zf.read(name)))
            for name in sorted(item for item in zf.namelist() if re.fullmatch(r"word/footer\d+\.xml", item)):
                parts_to_parse.append((name, "footer", zf.read(name)))

            all_parts = []
            for part_name, part_kind, xml_bytes in parts_to_parse:
                paragraphs = parse_part(
                    part_name,
                    xml_bytes,
                    part_kind,
                    style_catalog,
                    style_cache,
                    paragraph_profiles,
                    run_profiles,
                    paragraph_examples,
                    run_examples,
                )
                all_parts.append(
                    {
                        "part": part_name,
                        "part_kind": part_kind,
                        "paragraph_count": len(paragraphs),
                        "non_empty_paragraph_count": sum(1 for item in paragraphs if item["text"].strip()),
                        "paragraphs": paragraphs,
                    }
                )
    finally:
        if temp_docx_path.exists():
            temp_docx_path.unlink()

    main_part = next(part for part in all_parts if part["part_kind"] == "document")
    style_list = sorted(style_catalog["styles"].values(), key=lambda item: ((item.get("type") or ""), (item.get("name") or "")))

    return {
        "generated_at": dt.datetime.now(dt.timezone.utc).astimezone().isoformat(),
        "source_document": str(source_path),
        "output_purpose": "用于 agent 生成与模板一致的论文 Word 文本格式要求",
        "analysis_method": [
            "使用 Word COM 直接读取原始 .doc 模板并提取页面设置、节信息和页眉页脚快照",
            "使用 Word COM 识别文本框等非正文流对象，补足封面主标题等版式信息",
            "临时转存为 .docx 后解析 OpenXML，提取段落级与字符级格式",
            "对段落格式和字符格式分别去重，生成可复用 profile",
        ],
        "document_stats": word_snapshot["document_stats"],
        "page_setup": word_snapshot["page_setup"],
        "header_footer_text_snapshot": word_snapshot["header_footer_text_snapshot"],
        "shape_textboxes_snapshot": word_snapshot["shape_textboxes"],
        "style_catalog": {
            "doc_defaults": style_catalog["doc_defaults"],
            "styles": style_list,
        },
        "format_summary": build_format_summary(main_part["paragraphs"]),
        "agent_format_requirements": analyze_document_structure(main_part["paragraphs"], all_parts, word_snapshot),
        "paragraph_profiles": summarize_profiles(paragraph_profiles, paragraph_examples),
        "run_profiles": summarize_profiles(run_profiles, run_examples),
        "parts": all_parts,
    }


if __name__ == "__main__":
    main()
