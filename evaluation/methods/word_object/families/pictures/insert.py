"""families / pictures / insert with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Any, Dict, List, Optional
import re
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_horizontal_line_signature(docx_path: Path) -> Dict[str, Any]:
    """Extract Word's legacy VML horizontal-rule semantics and anchor."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    o = "urn:schemas-microsoft-com:office:office"
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    paragraphs = []
    line_records = []
    body = root.find("w:body", ns)
    body_paragraphs = [] if body is None else body.findall("./w:p", ns)
    for paragraph_index, paragraph in enumerate(body_paragraphs, start=1):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", ns))
        picts = paragraph.findall(".//w:pict", ns)
        drawings = paragraph.findall(".//w:drawing", ns)
        paragraph_line_count = 0
        other_object_count = len(drawings)
        for pict in picts:
            hr_elements = [
                node
                for node in pict.iter()
                if (node.get(f"{{{o}}}hr") or "").lower() in {"t", "true", "1"}
            ]
            if not hr_elements:
                other_object_count += 1
                continue
            paragraph_line_count += len(hr_elements)
            for node in hr_elements:
                style = node.get("style", "")
                height_match = re.search(
                    r"(?:^|;)\s*height\s*:\s*([0-9.]+)pt(?:;|$)",
                    style,
                    flags=re.IGNORECASE,
                )
                line_records.append(
                    {
                        "paragraph_index": paragraph_index,
                        "element": node.tag.rsplit("}", 1)[-1],
                        "height_points": (float(height_match.group(1)) if height_match else None),
                        "standard_rule": node.get(f"{{{o}}}hrstd"),
                        "alignment": node.get(f"{{{o}}}hralign"),
                        "percent_width": node.get(f"{{{o}}}hrpct"),
                        "fill_color": node.get("fillcolor"),
                        "stroked": node.get("stroked"),
                    }
                )
        paragraphs.append(
            {
                "text": text,
                "horizontal_line_count": paragraph_line_count,
                "other_object_count": other_object_count,
            }
        )
    section_layout = []
    for section in root.findall(".//w:sectPr", ns):
        page_size = section.find("w:pgSz", ns)
        margins = section.find("w:pgMar", ns)
        columns = section.find("w:cols", ns)

        def integer(element: Optional[ET.Element], name: str) -> Optional[int]:
            if element is None:
                return None
            value = element.get(f"{{{w}}}{name}")
            try:
                return int(value) if value is not None else None
            except ValueError:
                return None

        section_layout.append(
            {
                "page_width": integer(page_size, "w"),
                "page_height": integer(page_size, "h"),
                "orientation": (page_size.get(f"{{{w}}}orient") if page_size is not None else None),
                "margin_top": integer(margins, "top"),
                "margin_right": integer(margins, "right"),
                "margin_bottom": integer(margins, "bottom"),
                "margin_left": integer(margins, "left"),
                "gutter": integer(margins, "gutter"),
                "header_distance": integer(margins, "header"),
                "footer_distance": integer(margins, "footer"),
                "column_count": integer(columns, "num"),
                "column_spacing": integer(columns, "space"),
            }
        )
    return {
        "horizontal_line_count": len(line_records),
        "lines": line_records,
        "paragraphs": paragraphs,
        "section_layout": section_layout,
    }


def verify_inline_shapes(
    pred_path: Path,
    gold_path: Path,
    *,
    com_primary_with_saved_guardrail,
    error_with_details,
    read_inline_shapes,
) -> VerificationResult:
    """Verify the real Word horizontal-line object and its anchor."""
    return com_primary_with_saved_guardrail(
        lambda: verify_inline_shapes_com(
            pred_path,
            gold_path,
            error_with_details=error_with_details,
            read_inline_shapes=read_inline_shapes,
        ),
        lambda: verify_inline_shapes_saved(
            pred_path, gold_path, error_with_details=error_with_details
        ),
    )


def verify_inline_shapes_com(
    pred_path: Path, gold_path: Path, *, error_with_details, read_inline_shapes
) -> VerificationResult:
    gold_info = read_inline_shapes(str(gold_path))
    pred_info = read_inline_shapes(str(pred_path))

    gold_lines = [shape for shape in gold_info.get("inline_shapes", []) if shape.get("type") == 6]
    if (
        len(gold_lines) != 1
        or gold_info.get("inline_shape_count") != 1
        or gold_info.get("shape_count") != 0
    ):
        return error_with_details(
            "Gold does not contain exactly one Word horizontal-line object",
            {"match_mode": "word_com_horizontal_line", "gold": gold_info},
        )

    pred_lines = [shape for shape in pred_info.get("inline_shapes", []) if shape.get("type") == 6]
    if (
        len(pred_lines) != 1
        or pred_info.get("inline_shape_count") != 1
        or pred_info.get("shape_count") != 0
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted object is not exactly one Word horizontal line",
            details={
                "match_mode": "word_com_horizontal_line",
                "pred": pred_info,
                "gold": gold_info,
            },
        )

    if float(pred_lines[0].get("height") or 0) <= 0:
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted horizontal line is degenerate/invisible",
            details={
                "match_mode": "word_com_horizontal_line",
                "pred_line": pred_lines[0],
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Word horizontal-line object identity is present",
        details={
            "match_mode": "word_com_horizontal_line",
            "pred": pred_info,
        },
    )


def verify_inline_shapes_saved(
    pred_path: Path, gold_path: Path, *, error_with_details
) -> VerificationResult:
    gold = extract_saved_horizontal_line_signature(gold_path)
    pred = extract_saved_horizontal_line_signature(pred_path)
    if gold.get("horizontal_line_count") != 1:
        return error_with_details(
            "Gold saved horizontal-line evidence is missing or ambiguous",
            {"match_mode": "saved_ooxml_horizontal_line", "gold": gold},
        )
    if pred.get("horizontal_line_count") != 1:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved output lacks a real Word horizontal rule",
            details={
                "match_mode": "saved_ooxml_horizontal_line",
                "gold": gold,
                "pred": pred,
            },
        )
    pred_line = pred["lines"][0]
    if pred_line.get("height_points") is None or pred_line.get("height_points") <= 0:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved Word horizontal rule is degenerate",
            details={
                "match_mode": "saved_ooxml_horizontal_line",
                "pred": pred,
            },
        )
    if pred.get("paragraphs") != gold.get("paragraphs"):
        return VerificationResult(
            passed=False,
            message="FAIL: Horizontal line anchor or surrounding text differs",
            details={
                "match_mode": "saved_ooxml_horizontal_line_anchor",
                "gold": gold,
                "pred": pred,
            },
        )

    def section_layouts_equivalent(
        left: List[Dict[str, Any]],
        right: List[Dict[str, Any]],
    ) -> bool:
        if len(left) != len(right):
            return False
        tolerant_twip_fields = {"header_distance", "footer_distance", "column_spacing"}
        for left_section, right_section in zip(left, right):
            for key in set(left_section) | set(right_section):
                left_value = left_section.get(key)
                right_value = right_section.get(key)
                if key in tolerant_twip_fields:
                    if left_value is None or right_value is None:
                        if left_value != right_value:
                            return False
                    elif abs(left_value - right_value) > 20:
                        return False
                elif left_value != right_value:
                    return False
        return True

    if not section_layouts_equivalent(
        pred.get("section_layout", []), gold.get("section_layout", [])
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Horizontal-line insertion changed section layout",
            details={
                "match_mode": "saved_ooxml_horizontal_line_section",
                "gold": gold,
                "pred": pred,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Saved Word horizontal-line identity and anchor match",
        details={
            "match_mode": "saved_ooxml_horizontal_line",
            "line": pred_line,
        },
    )
