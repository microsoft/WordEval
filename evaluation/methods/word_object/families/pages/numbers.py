"""families / pages / numbers with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re
import unicodedata
from ...observations import ObservedProperties, require_observations


def verify_page_number_field(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    get_extract_body_page_number_layout,
    read_document_fields,
    read_document_text,
) -> VerificationResult:
    """Verify a current-position page-number building block in the body."""

    def _field_keyword(field: Dict) -> str:
        code = " ".join((field.get("code") or "").upper().split())
        if field.get("type") == 26 or re.search(r"\bNUMPAGES\b", code):
            return "NUMPAGES"
        if field.get("type") == 33 or re.search(r"\bPAGE\b", code):
            return "PAGE"
        return ""

    def _page_fields(path: Path) -> List[Dict]:
        fields = read_document_fields(str(path), update_fields=True)
        return [field for field in fields if _field_keyword(field)]

    def _masked_text(text: str) -> str:
        text = unicodedata.normalize("NFC", text or "").replace("\xa0", " ")
        return text.rstrip("\r\n")

    def _font_signature(font: Optional[Dict]) -> Dict:
        font = font or {}

        def _effective_boolean(value: Any) -> Any:
            if isinstance(value, list):
                return [(start, end, _effective_boolean(item)) for start, end, item in value]
            return value not in (None, False, 0)

        def _effective_number(value: Any) -> Any:
            if isinstance(value, list):
                return [(start, end, _effective_number(item)) for start, end, item in value]
            try:
                return round(float(value or 0.0), 2)
            except (TypeError, ValueError):
                return 0.0

        return {
            "name": font.get("name"),
            "size": _effective_number(font.get("size")),
            "bold": _effective_boolean(font.get("bold")),
            "italic": _effective_boolean(font.get("italic")),
            "color": font.get("color"),
            "spacing": _effective_number(font.get("spacing")),
        }

    def _normalized_layout(layout: Dict) -> Dict:
        paragraphs = []
        for paragraph in layout.get("paragraphs") or []:
            units = []
            for unit in paragraph.get("units") or []:
                units.append(
                    {
                        "kind": unit.get("kind"),
                        "text": unit.get("text"),
                        "keyword": unit.get("keyword"),
                        "font": _font_signature(unit.get("font")),
                    }
                )
            paragraphs.append(
                {
                    "paragraph_index": paragraph.get("paragraph_index"),
                    "template": paragraph.get("template"),
                    "units": units,
                    "alignment": paragraph.get("alignment"),
                    "left_indent": paragraph.get("left_indent"),
                    "right_indent": paragraph.get("right_indent"),
                    "first_line_indent": paragraph.get("first_line_indent"),
                }
            )
        return {
            "document_template": _masked_text(layout.get("document_template", layout.get("document_text", ""))),
            "page_count": layout.get("page_count"),
            "paragraphs": paragraphs,
            "fields": [
                {
                    "paragraph_index": field.get("paragraph_index"),
                    "keyword": field.get("keyword"),
                    "font": _font_signature(field.get("font")),
                }
                for field in (layout.get("fields") or [])
            ],
        }

    gold_page_fields = _page_fields(gold_path)
    pred_page_fields = _page_fields(pred_path)
    if not gold_page_fields:
        return policies_content_operations.error_with_details(
            "Gold contains no dynamic PAGE/NUMPAGES field target",
            {"gold_fields": gold_page_fields, "spec": spec or {}},
        )
    if not pred_page_fields:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: No dynamic PAGE/NUMPAGES field found in predicted "
                "document (plain text is not accepted)"
            ),
            details={
                "gold_page_fields": gold_page_fields,
                "pred_page_fields": pred_page_fields,
            },
        )

    gold_keywords = [_field_keyword(field) for field in gold_page_fields]
    pred_keywords = [_field_keyword(field) for field in pred_page_fields]
    if pred_keywords != gold_keywords:
        return VerificationResult(
            passed=False,
            message="FAIL: Dynamic page-field sequence differs",
            details={
                "match_mode": "body_page_number_field_layout",
                "gold_field_keywords": gold_keywords,
                "pred_field_keywords": pred_keywords,
                "gold_page_fields": gold_page_fields,
                "pred_page_fields": pred_page_fields,
            },
        )

    layout_extractor = get_extract_body_page_number_layout()
    if layout_extractor is not None:
        gold_layout = _normalized_layout(layout_extractor(str(gold_path)))
        pred_layout = _normalized_layout(layout_extractor(str(pred_path)))
    else:
        # Compatibility for focused fake-core tests. Production WOM always
        # supplies the richer extractor above.
        gold_layout = {
            "document_template": _masked_text(read_document_text(str(gold_path))),
            "fields": [{"keyword": value} for value in gold_keywords],
        }
        pred_layout = {
            "document_template": _masked_text(read_document_text(str(pred_path))),
            "fields": [{"keyword": value} for value in pred_keywords],
        }

    if not gold_layout.get("paragraphs", gold_layout.get("fields")):
        return policies_content_operations.error_with_details(
            "Gold body page-number layout could not be established",
            {"gold_layout": gold_layout, "spec": spec or {}},
        )
    if pred_layout != gold_layout:
        return VerificationResult(
            passed=False,
            message="FAIL: Body page-number preset layout differs",
            details={
                "match_mode": "body_page_number_field_layout",
                "gold": gold_layout,
                "pred": pred_layout,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Page number field layout matches",
        details={
            "match_mode": "body_page_number_field_layout",
            "field_keywords": pred_keywords,
            "layout": pred_layout,
        },
    )


def verify_margin_page_numbers(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    get_optional_reader,
    read_document_text,
    read_shape_text_styles,
) -> VerificationResult:
    """
    Verify margin page-number gallery insertion via shape text style signatures.

    These variants are typically represented as shapes/textboxes in the main
    document story, so shape-level text style comparison is more reliable than
    body paragraph style checks.
    """

    def _visible_template(text: str, mask_digits: bool = False) -> str:
        normalized = unicodedata.normalize("NFC", text or "").replace("\xa0", " ")
        if mask_digits:
            normalized = re.sub(r"\d+", "#", normalized)
        return normalized.rstrip("\r\n")

    def _font_signature(font: Optional[Dict]) -> Dict:
        font = font or {}
        return {
            key: font.get(key)
            for key in (
                "name",
                "size",
                "bold",
                "italic",
                "color",
                "color_index",
                "spacing",
            )
        }

    def _shape_sig(shape_info: Dict) -> Dict:
        page_field_count = shape_info.get("page_field_count", 0) or 0
        num_pages_field_count = shape_info.get("num_pages_field_count", 0) or 0
        has_dynamic_page_field = page_field_count > 0 or num_pages_field_count > 0
        return {
            "type": shape_info.get("type"),
            "auto_shape_type": shape_info.get("auto_shape_type"),
            "relative_horizontal_position": shape_info.get("relative_horizontal_position"),
            "relative_vertical_position": shape_info.get("relative_vertical_position"),
            "left": shape_info.get("left"),
            "top": shape_info.get("top"),
            "width": shape_info.get("width"),
            "height": shape_info.get("height"),
            "rotation": shape_info.get("rotation"),
            "wrap_type": shape_info.get("wrap_type"),
            "text_orientation": shape_info.get("text_orientation"),
            "text_margins": [
                shape_info.get("text_margin_left"),
                shape_info.get("text_margin_right"),
                shape_info.get("text_margin_top"),
                shape_info.get("text_margin_bottom"),
            ],
            "fill": {
                "visible": shape_info.get("fill_visible"),
                "type": shape_info.get("fill_type"),
                "color": shape_info.get("fill_color"),
                "theme_color": shape_info.get("fill_theme_color"),
                "transparency": shape_info.get("fill_transparency"),
            },
            "line": {
                "visible": shape_info.get("line_visible"),
                "style": shape_info.get("line_style"),
                "dash_style": shape_info.get("line_dash_style"),
                "weight": shape_info.get("line_weight"),
                "color": shape_info.get("line_color"),
                "theme_color": shape_info.get("line_theme_color"),
                "transparency": shape_info.get("line_transparency"),
            },
            "paragraph_styles": [
                (_s or "").strip() for _s in (shape_info.get("paragraph_styles") or [])
            ],
            "paragraph_texts": [
                _visible_template(t, mask_digits=has_dynamic_page_field and "paragraph_templates" not in shape_info)
                for t in (shape_info.get("paragraph_templates", shape_info.get("paragraph_texts")) or [])
            ],
            "paragraph_alignments": shape_info.get("paragraph_alignments") or [],
            "paragraph_fonts": [
                _font_signature(font) for font in (shape_info.get("paragraph_fonts") or [])
            ],
            "field_keywords": sorted([k for k in (shape_info.get("field_keywords") or []) if k]),
            "field_fonts": [
                _font_signature(font) for font in (shape_info.get("field_fonts") or [])
            ],
            "page_field_count": page_field_count,
            "num_pages_field_count": num_pages_field_count,
        }

    def _shape_sig_key(sig: Dict):
        return (
            str(sig.get("type")),
            str(sig.get("auto_shape_type")),
            str(sig.get("relative_horizontal_position")),
            str(sig.get("relative_vertical_position")),
            tuple(sig.get("paragraph_styles", [])),
            tuple(sig.get("paragraph_texts", [])),
            tuple(sig.get("field_keywords", [])),
            sig.get("page_field_count", 0),
            sig.get("num_pages_field_count", 0),
        )

    def _shape_sig_list(shapes: List[Dict]) -> List[Dict]:
        body_sigs = []
        header_footer_entries = []
        for sh in shapes:
            sig = _shape_sig(sh)
            if sh.get("story") == "body":
                body_sigs.append(sig)
            else:
                header_footer_entries.append((sig, sh.get("story")))

        # Word can expose the same anchored margin page-number shape through
        # multiple shape collections. Keep same-story duplicates visible, but
        # avoid double-counting cross-story COM exposure of one logical shape.
        body_keys = {_shape_sig_key(sig) for sig in body_sigs}
        sigs = list(body_sigs)
        seen_header_footer_stories = {}
        for sig, story in header_footer_entries:
            key = _shape_sig_key(sig)
            if key in body_keys:
                continue
            seen_stories = seen_header_footer_stories.setdefault(key, set())
            if seen_stories and story not in seen_stories:
                continue
            sigs.append(sig)
            seen_stories.add(story)
        return sorted(
            sigs,
            key=lambda s: (
                str(s.get("type")),
                str(s.get("auto_shape_type")),
                str(s.get("relative_horizontal_position")),
                str(s.get("paragraph_styles", [])),
                str(s.get("paragraph_texts", [])),
                str(s.get("field_keywords", [])),
                s.get("page_field_count", 0),
                s.get("num_pages_field_count", 0),
            ),
        )

    observations = {owner: read_shape_text_styles(str(path), include_header_footer=True)
                    for owner, path in (("gold", gold_path), ("pred", pred_path))}
    from ..shapes.margin_styles import validate_margin_shape_styles

    def validate(owner):
        error = validate_margin_shape_styles({owner: observations[owner]})
        if error is not None:
            return error
        for shape in observations[owner]:
            if isinstance(shape, ObservedProperties):
                error = require_observations({owner: shape}, {"paragraph_templates"},
                                             context="margin page-number field spans")
                if error is not None:
                    return error

    error = validate("gold")
    if error is not None:
        return error
    gold_sigs = _shape_sig_list(observations["gold"])

    if not gold_sigs or not any(
        signature.get("page_field_count", 0) > 0 for signature in gold_sigs
    ):
        return policies_content_operations.error_with_details(
            "Gold contains no dynamic margin PAGE field shape",
            {"gold": gold_sigs, "spec": spec or {}},
        )

    error = validate("pred")
    if error is not None:
        # A known identity mismatch cannot be rescued by an unreadable item.
        # Use only fields in the existing cross-story identity key: geometry
        # alone could differ on an exposure that the matcher deduplicates.
        identity_fields = {
            "type": "type", "auto_shape_type": "auto_shape_type",
            "relative_horizontal_position": "relative_horizontal_position",
            "relative_vertical_position": "relative_vertical_position",
            "paragraph_styles": "paragraph_styles", "paragraph_templates": "paragraph_texts",
            "field_keywords": "field_keywords", "page_field_count": "page_field_count",
            "num_pages_field_count": "num_pages_field_count",
        }
        for shape in observations["pred"]:
            known = {field: projected for field, projected in identity_fields.items()
                     if field in shape and require_observations({"pred": shape}, {field},
                        context="margin shape identity") is None}
            if not known:
                continue
            signature = _shape_sig(shape)
            candidates = [[dict(property=field, gold=gold[projected], pred=signature[projected])
                           for field, projected in known.items() if signature[projected] != gold[projected]]
                          for gold in gold_sigs]
            if all(candidates):
                return VerificationResult(False, "FAIL: Observed margin shape cannot match any reference shape",
                    details={"match_mode": "observed_margin_shape_identity_mismatch",
                             "candidate_differences": candidates, "shape_index": shape.get("index"),
                             "incomplete_observation": error.to_dict()})
        return error
    pred_sigs = _shape_sig_list(observations["pred"])

    if len(gold_sigs) != len(pred_sigs):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Margin page number shape count mismatch (pred={len(pred_sigs)}, gold={len(gold_sigs)})",
            details={
                "match_mode": "margin_page_number_shape_signature",
                "pred_count": len(pred_sigs),
                "gold_count": len(gold_sigs),
                "pred": pred_sigs,
                "gold": gold_sigs,
            },
        )

    def _differences(
        pred_value: Any,
        gold_value: Any,
        path: str = "",
    ) -> List[Dict]:
        if isinstance(gold_value, dict) and isinstance(pred_value, dict):
            differences = []
            for key in sorted(set(gold_value) | set(pred_value)):
                differences.extend(
                    _differences(
                        pred_value.get(key),
                        gold_value.get(key),
                        f"{path}.{key}" if path else key,
                    )
                )
            return differences
        if isinstance(gold_value, list) and isinstance(pred_value, list):
            if len(gold_value) != len(pred_value):
                return [
                    {
                        "path": path,
                        "pred": pred_value,
                        "gold": gold_value,
                    }
                ]
            differences = []
            for index, (pred_item, gold_item) in enumerate(zip(pred_value, gold_value)):
                differences.extend(
                    _differences(
                        pred_item,
                        gold_item,
                        f"{path}[{index}]",
                    )
                )
            return differences
        if (
            isinstance(gold_value, (int, float))
            and not isinstance(gold_value, bool)
            and isinstance(pred_value, (int, float))
            and not isinstance(pred_value, bool)
            and abs(float(pred_value) - float(gold_value)) <= 0.5
        ):
            return []
        if pred_value == gold_value:
            return []
        return [{"path": path, "pred": pred_value, "gold": gold_value}]

    shape_differences = _differences(pred_sigs, gold_sigs)
    if shape_differences:
        return VerificationResult(
            passed=False,
            message="FAIL: Margin page number styles differ",
            details={
                "match_mode": "margin_page_number_shape_signature",
                "pred": pred_sigs,
                "gold": gold_sigs,
                "differences": shape_differences,
            },
        )

    preservation_differences = {}
    gold_text = read_document_text(str(gold_path))
    pred_text = read_document_text(str(pred_path))
    if pred_text != gold_text:
        preservation_differences["body_text"] = {
            "pred": pred_text,
            "gold": gold_text,
        }
    for label, extractor_name in (
        ("page_info", "extract_page_info"),
        ("section_properties", "extract_section_properties"),
    ):
        extractor = get_optional_reader(extractor_name)
        if extractor is None:
            continue
        gold_value = extractor(str(gold_path))
        pred_value = extractor(str(pred_path))
        if gold_value is None:
            return policies_content_operations.error_with_details(
                f"Gold margin-page-number {label} evidence could not be extracted",
                {"gold_shapes": gold_sigs, "spec": spec or {}},
            )
        if pred_value != gold_value:
            preservation_differences[label] = {
                "pred": pred_value,
                "gold": gold_value,
            }
    if preservation_differences:
        return VerificationResult(
            passed=False,
            message=("FAIL: Margin page number insertion changed unrelated " "document layout"),
            details={
                "match_mode": "margin_page_number_visible_layout",
                "differences": preservation_differences,
                "shape_count": len(gold_sigs),
            },
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Margin page number styles match ({len(gold_sigs)} shapes)",
        details={
            "match_mode": "margin_page_number_visible_layout",
            "shape_count": len(gold_sigs),
        },
    )
