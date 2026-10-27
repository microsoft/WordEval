"""families / headers / verify with explicit evidence dependencies."""

from datetime import datetime
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.documents import verify as families_documents_verify
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re
from ...observations import require_observations, ObservedProperties


def verify_header_footer(
    pred_path: Path,
    gold_path: Path,
    part_type: str = "footer",
    check_type: str = "content_match",
    spec: Optional[Dict] = None,
    *,
    verify_body_preservation: bool = True,
    read_document_text,
    read_header_footer_content,
) -> VerificationResult:
    """
    Verify header/footer content.

    Args:
        part_type: 'header' or 'footer'
        check_type: 'content_match', 'empty'/'not_exists', or 'not_empty'/'exists'
    """
    gold_content = read_header_footer_content(str(gold_path))
    pred_content = read_header_footer_content(str(pred_path))

    if gold_content is None:
        return policies_content_operations.error_with_details(
            "Could not extract gold header/footer content"
        )
    if pred_content is None:
        return policies_content_operations.error_with_details(
            "Could not extract predicted header/footer content"
        )

    if verify_body_preservation:
        body_result = families_documents_verify.verify_body_text_matches(
            pred_path, gold_path, read_document_text=read_document_text
        )
        if not body_result.passed:
            return body_result

    gold_parts = gold_content["headers" if part_type == "header" else "footers"]
    pred_parts = pred_content["headers" if part_type == "header" else "footers"]

    # Normalize check_type aliases
    if check_type == "not_exists":
        check_type = "empty"
    elif check_type == "exists":
        check_type = "not_empty"

    required = {"type", "text", "field_count", "inline_shape_count", "shape_count", "content_control_values"}
    if check_type == "no_shape":
        required = {"type", "inline_shape_count", "shape_count"}
    elif check_type == "page_field_match":
        required = {"type", "page_field_count", "num_pages_field_count"}
    elif check_type == "style_match":
        required |= {"text_template", "paragraph_signatures", "field_codes", "page_field_count", "num_pages_field_count"}
    for owner, parts in (("gold", gold_parts), ("pred", pred_parts)):
        for index, part in enumerate(parts, 1):
            # Historical injected projections may omit unexercised fields;
            # native readers always expose the complete projection.
            fields = required if isinstance(part, ObservedProperties) else required & part.keys()
            error = require_observations({owner: part}, fields, context=f"{part_type} {index} {check_type}")
            if error is not None:
                return error
            if check_type == "style_match":
                for paragraph in part.get("paragraph_signatures", []):
                    fields = {"style", "alignment", "font_size", "font_color", "bold", "italic"}
                    for edge in ("top", "left", "bottom", "right"):
                        fields.add(f"{edge}_border_line_style")
                        if paragraph.get(f"{edge}_border_line_style") not in (None, 0):
                            fields |= {f"{edge}_border_line_width", f"{edge}_border_color"}
                    if not isinstance(paragraph, ObservedProperties):
                        fields &= paragraph.keys()
                    error = require_observations({owner: paragraph}, fields, context=f"{part_type} paragraph style")
                    if error is not None:
                        return error

    def _parse_date_value(value: str):
        text = re.sub(r"\s+", " ", (value or "").strip())
        if not text:
            return None
        for fmt in ("%Y-%m-%d", "%B %d, %Y", "%b %d, %Y", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                pass
        return None

    def _content_control_values_equal(key: str, pred_val: str, gold_val: str) -> bool:
        if pred_val == gold_val:
            return True
        if "date" in (key or "").lower():
            pred_date = _parse_date_value(pred_val)
            gold_date = _parse_date_value(gold_val)
            if pred_date is not None and gold_date is not None:
                return pred_date == gold_date
        return False

    def _semantic_part_payload(part: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "type": part.get("type"),
            "text": (part.get("text") or "").strip(),
            "field_count": int(part.get("field_count") or 0),
            "page_field_count": int(part.get("page_field_count") or 0),
            "num_pages_field_count": int(part.get("num_pages_field_count") or 0),
            "inline_shape_count": int(part.get("inline_shape_count") or 0),
            "shape_count": int(part.get("shape_count") or 0),
            "content_control_values": part.get("content_control_values", {}) or {},
        }

    def _has_semantic_part_content(part: Dict[str, Any]) -> bool:
        payload = _semantic_part_payload(part)
        return bool(
            payload["text"]
            or payload["field_count"]
            or payload["page_field_count"]
            or payload["num_pages_field_count"]
            or payload["inline_shape_count"]
            or payload["shape_count"]
            or payload["content_control_values"]
        )

    if check_type == "empty":
        pred_payloads = [
            _semantic_part_payload(part) for part in pred_parts if _has_semantic_part_content(part)
        ]
        if pred_payloads:
            return VerificationResult(
                passed=False,
                message=f"FAIL: {part_type} should be empty but has content",
                details={"pred_content": pred_payloads},
            )
        return VerificationResult(passed=True, message=f"PASS: {part_type} is empty", details={})

    elif check_type == "not_empty":
        pred_payloads = [
            _semantic_part_payload(part) for part in pred_parts if _has_semantic_part_content(part)
        ]
        if not pred_payloads:
            return VerificationResult(
                passed=False,
                message=f"FAIL: {part_type} should have content but is empty",
                details={},
            )
        return VerificationResult(
            passed=True,
            message=f"PASS: {part_type} has content",
            details={"content": pred_payloads},
        )

    elif check_type == "no_shape":

        def _shape_signature(parts: List[Dict[str, Any]]) -> List[Dict[str, int]]:
            return sorted(
                [
                    {
                        "type": int(part.get("type") or 0),
                        "inline_shape_count": int(part.get("inline_shape_count") or 0),
                        "shape_count": int(part.get("shape_count") or 0),
                    }
                    for part in parts
                ],
                key=lambda item: (
                    item["type"],
                    item["inline_shape_count"],
                    item["shape_count"],
                ),
            )

        gold_signature = _shape_signature(gold_parts)
        pred_signature = _shape_signature(pred_parts)
        if pred_signature != gold_signature:
            return VerificationResult(
                passed=False,
                message=f"FAIL: {part_type} shape content differs",
                details={
                    "match_mode": "shape_count_match",
                    "gold": gold_signature,
                    "pred": pred_signature,
                },
            )
        return VerificationResult(
            passed=True,
            message=f"PASS: {part_type} shape content matches",
            details={
                "match_mode": "shape_count_match",
                "parts": gold_signature,
            },
        )

    elif check_type == "page_field_match":
        # Composition instructions request a page number in a header or
        # footer without naming an exact gallery alignment. Preserve the
        # semantic contract (a real PAGE/NUMPAGES field in the same story
        # type) without requiring an unstated left/center/right style.
        def _page_field_signature(parts: List[Dict[str, Any]]) -> List[Dict[str, int]]:
            return sorted(
                [
                    {
                        "type": int(part.get("type") or 0),
                        "page_field_count": int(part.get("page_field_count") or 0),
                        "num_pages_field_count": int(part.get("num_pages_field_count") or 0),
                    }
                    for part in parts
                    if (part.get("page_field_count") or 0)
                    or (part.get("num_pages_field_count") or 0)
                ],
                key=lambda item: (
                    item["type"],
                    item["page_field_count"],
                    item["num_pages_field_count"],
                ),
            )

        gold_signature = _page_field_signature(gold_parts)
        pred_signature = _page_field_signature(pred_parts)
        if not gold_signature:
            return policies_content_operations.error_with_details(
                f"Gold {part_type} page-number field evidence is missing",
                {
                    "match_mode": "page_field_match",
                    "gold": gold_signature,
                },
            )
        if gold_signature != pred_signature:
            return VerificationResult(
                passed=False,
                message=f"FAIL: {part_type} page-number fields differ",
                details={
                    "match_mode": "page_field_match",
                    "gold": gold_signature,
                    "pred": pred_signature,
                },
            )
        return VerificationResult(
            passed=True,
            message=f"PASS: {part_type} page-number fields match",
            details={
                "match_mode": "page_field_match",
                "parts": gold_signature,
            },
        )

    elif check_type == "style_match":
        idMso = (spec or {}).get("idMso", "")
        compare_page_text_skeleton = True

        def _normalize_story_paragraph_style(style_name: Optional[str]) -> Optional[str]:
            style = (style_name or "").strip()
            if not style:
                return style_name

            normalized = style.lower()
            if part_type == "header" and normalized in {"normal", "header", "footer"}:
                return "__header_story_base__"
            if part_type == "footer" and normalized in {"normal", "header", "footer"}:
                return "__footer_story_base__"
            return style_name

        def _normalize_field_code(code: str) -> str:
            # Ignore whitespace noise and common field formatting switches.
            c = " ".join((code or "").upper().split())
            c = c.replace("\\* MERGEFORMAT", "").strip()
            return c

        def _normalize_page_text_skeleton(text: str) -> str:
            # Keep literal decorations (e.g., brackets/pipes/words), ignore dynamic numbers.
            t = (text or "").replace("\r", " ").replace("\x07", " ")
            t = re.sub(r"\d+", "#", t)
            t = " ".join(t.split())
            return t.strip()

        def _is_placeholder_value(value: str) -> bool:
            v = (value or "").strip()
            return v.startswith("[") and v.endswith("]")

        def _has_only_placeholder_controls(part: Dict) -> bool:
            cc_values = part.get("content_control_values", {}) or {}
            return bool(cc_values) and all(
                _is_placeholder_value(value) for value in cc_values.values()
            )

        def _field_keyword(code: str) -> str:
            c = _normalize_field_code(code)
            # Explicit page-number field keywords first (what we care about here)
            if "NUMPAGES" in c:
                return "NUMPAGES"
            if "PAGE" in c:
                return "PAGE"
            # Fallback: first token for other field types
            return c.split(" ")[0] if c else ""

        def _part_signature(part: Dict) -> Dict:
            raw_para_sigs = list(part.get("paragraph_signatures", []) or [])
            # Word often adds an extra trailing empty header/footer paragraph in gold only.
            # Trim it so style-match focuses on meaningful content/style structure.
            while raw_para_sigs and not (raw_para_sigs[-1].get("text") or "").strip():
                raw_para_sigs.pop()

            para_sigs = []
            for para in raw_para_sigs:
                normalized_style = _normalize_story_paragraph_style(para.get("style"))
                border_signatures = {}
                for border_name in ("top", "left", "bottom", "right"):
                    line_style = para.get(f"{border_name}_border_line_style")
                    border_signature = {"line_style": line_style}
                    # Disabled borders have no visual width or color.  Ignore
                    # COM/default noise for those properties, while retaining
                    # the render-relevant attributes for an enabled edge.
                    if line_style not in (None, 0):
                        border_signature.update(
                            {
                                "line_width": para.get(f"{border_name}_border_line_width"),
                                "color": para.get(f"{border_name}_border_color"),
                            }
                        )
                    border_signatures[border_name] = border_signature
                para_sigs.append(
                    {
                        "style": normalized_style,
                        "alignment": para.get("alignment"),
                        # Built-in gallery text inherits the document theme.
                        # Comparing Word's resolved font name makes an otherwise
                        # identical header/footer fail across Office locales
                        # (for example Aptos vs Malgun Gothic).  Font-family
                        # commands are verified by their dedicated formatting
                        # contracts, not by this gallery-layout matcher.
                        "font_size": para.get("font_size"),
                        "font_color": para.get("font_color"),
                        "bold": para.get("bold"),
                        "italic": para.get("italic"),
                        "border_signatures": border_signatures,
                    }
                )

            normalized_codes = sorted(
                [_normalize_field_code(c) for c in (part.get("field_codes", []) or [])]
            )
            field_keywords = sorted([k for k in (_field_keyword(c) for c in normalized_codes) if k])
            has_page_fields = (part.get("page_field_count", 0) or 0) > 0 or (
                part.get("num_pages_field_count", 0) or 0
            ) > 0
            page_text_skeleton = None
            if compare_page_text_skeleton and has_page_fields:
                page_text_skeleton = (
                    " ".join(part["text_template"].replace("\r", " ").replace("\x07", " ").split())
                    if "text_template" in part else _normalize_page_text_skeleton(part.get("text", ""))
                )
            semantic_cc_count = (
                0
                if _has_only_placeholder_controls(part)
                else len(part.get("content_control_values", {}) or {})
            )
            trimmed_para_styles = [
                normalized_style
                for normalized_style in (
                    _normalize_story_paragraph_style(p.get("style")) for p in raw_para_sigs
                )
                if normalized_style is not None
            ]

            return {
                "type": part.get("type"),
                "paragraph_count": len(raw_para_sigs),
                "paragraph_styles": trimmed_para_styles,
                "paragraph_signatures": para_sigs,
                "field_keywords": field_keywords,
                "page_field_count": part.get("page_field_count", 0),
                "num_pages_field_count": part.get("num_pages_field_count", 0),
                "page_text_skeleton": page_text_skeleton,
                "inline_shape_count": part.get("inline_shape_count", 0),
                "shape_count": part.get("shape_count", 0),
                "content_control_count": semantic_cc_count,
            }

        def _sort_key(sig: Dict):
            return (
                sig.get("type", 0),
                str(sig.get("paragraph_styles", [])),
                str(sig.get("field_keywords", [])),
                sig.get("page_field_count", 0),
                sig.get("num_pages_field_count", 0),
                str(sig.get("page_text_skeleton") or ""),
                sig.get("shape_count", 0),
                sig.get("inline_shape_count", 0),
                sig.get("content_control_count", 0),
            )

        gold_sigs = sorted([_part_signature(p) for p in gold_parts], key=_sort_key)
        pred_sigs = sorted([_part_signature(p) for p in pred_parts], key=_sort_key)

        if not gold_sigs:
            return policies_content_operations.error_with_details(
                f"Gold {part_type} style evidence is missing",
                {
                    "match_mode": "style_match",
                    "gold": gold_sigs,
                },
            )

        if gold_sigs == pred_sigs:
            # Style matches — also compare CC fill values
            gold_cc = {}
            pred_cc = {}
            for p in gold_parts:
                gold_cc.update(p.get("content_control_values", {}))
            for p in pred_parts:
                pred_cc.update(p.get("content_control_values", {}))

            cc_mismatches = []
            for key, gold_val in gold_cc.items():
                if _is_placeholder_value(gold_val):
                    continue
                pred_val = pred_cc.get(key, "")
                if gold_val and not _content_control_values_equal(key, pred_val, gold_val):
                    cc_mismatches.append({"field": key, "gold": gold_val, "pred": pred_val})

            if cc_mismatches:
                return VerificationResult(
                    passed=False,
                    message=f"FAIL: {part_type} content control values differ ({len(cc_mismatches)} field(s))",
                    details={
                        "match_mode": "style_match",
                        "mismatches": cc_mismatches,
                        "gold_cc": gold_cc,
                        "pred_cc": pred_cc,
                    },
                )

            placeholder_only_gold = bool(gold_cc) and all(
                _is_placeholder_value(value) for value in gold_cc.values()
            )
            if placeholder_only_gold:

                def _placeholder_tokens(parts):
                    tokens = []
                    for part in parts:
                        text = (part.get("text") or "").replace("\r", " ").replace("\x07", " ")
                        tokens.extend(re.findall(r"\[[^\[\]]+\]", text))
                    return sorted(token.strip() for token in tokens)

                gold_tokens = _placeholder_tokens(gold_parts)
                pred_tokens = _placeholder_tokens(pred_parts)
                if gold_tokens != pred_tokens:
                    return VerificationResult(
                        passed=False,
                        message=f"FAIL: {part_type} placeholder text differs",
                        details={
                            "match_mode": "style_match",
                            "gold_placeholders": gold_tokens,
                            "pred_placeholders": pred_tokens,
                        },
                    )

            # Fallback: when no CCs exist (e.g. Blank galleries where CCs
            # dissolve into plain text runs), compare the part text directly.
            # PAGE/NUMPAGES fields are intentionally excluded: their visible
            # numbers are dynamic pagination results and can be stale or
            # environment-dependent under Word COM. The field keyword,
            # placement, style, and literal decoration skeleton were already
            # compared in the signature above.
            has_dynamic_page_fields = compare_page_text_skeleton and any(
                sig.get("page_field_count", 0) or sig.get("num_pages_field_count", 0)
                for sig in gold_sigs + pred_sigs
            )
            if not gold_cc and not has_dynamic_page_fields:

                def _norm_part_text(parts):
                    return sorted(re.sub(r"\s+", " ", (p.get("text") or "")).strip() for p in parts)

                gold_texts = _norm_part_text(gold_parts)
                pred_texts = _norm_part_text(pred_parts)
                if gold_texts != pred_texts:
                    return VerificationResult(
                        passed=False,
                        message=f"FAIL: {part_type} text content differs",
                        details={
                            "match_mode": "style_match",
                            "gold_texts": gold_texts,
                            "pred_texts": pred_texts,
                        },
                    )

            return VerificationResult(
                passed=True,
                message=f"PASS: {part_type} style matches",
                details={"match_mode": "style_match", "parts": gold_sigs, "cc_values": gold_cc},
            )

        # Provide a compact diff-oriented payload to aid debugging
        return VerificationResult(
            passed=False,
            message=f"FAIL: {part_type} style differs",
            details={
                "match_mode": "style_match",
                "gold_count": len(gold_sigs),
                "pred_count": len(pred_sigs),
                "gold": gold_sigs,
                "pred": pred_sigs,
            },
        )

    else:  # content_match
        gold_texts = sorted([p["text"] for p in gold_parts])
        pred_texts = sorted([p["text"] for p in pred_parts])

        if gold_texts != pred_texts:
            return VerificationResult(
                passed=False,
                message=f"FAIL: {part_type} content differs",
                details={"gold": gold_texts, "pred": pred_texts},
            )

        # Also compare content control fill values (inside shapes etc.)
        gold_cc = {}
        pred_cc = {}
        for p in gold_parts:
            gold_cc.update(p.get("content_control_values", {}))
        for p in pred_parts:
            pred_cc.update(p.get("content_control_values", {}))

        cc_mismatches = []
        for key, gold_val in gold_cc.items():
            if gold_val.startswith("[") and gold_val.endswith("]"):
                continue
            pred_val = pred_cc.get(key, "")
            if gold_val and not _content_control_values_equal(key, pred_val, gold_val):
                cc_mismatches.append({"field": key, "gold": gold_val, "pred": pred_val})

        if cc_mismatches:
            return VerificationResult(
                passed=False,
                message=f"FAIL: {part_type} content control values differ ({len(cc_mismatches)} field(s))",
                details={"mismatches": cc_mismatches, "gold_cc": gold_cc, "pred_cc": pred_cc},
            )

        return VerificationResult(
            passed=True,
            message=f"PASS: {part_type} content matches",
            details={"content": gold_texts, "cc_values": gold_cc},
        )


def verify_header_footer_composition(
    pred_path: Path, gold_path: Path, *, read_header_footer_content
) -> VerificationResult:
    """Verify header/footer text plus semantic fields and shapes.

    Composition cases include page-number insertion and watermark removal,
    so text alone is not enough: literal "1" text must not replace PAGE
    fields, and watermark shapes must be removed.
    """
    gold_hf = read_header_footer_content(str(gold_path))
    pred_hf = read_header_footer_content(str(pred_path))

    if gold_hf is None:
        return VerificationResult.error_result("Gold: Failed to extract header/footer")
    if pred_hf is None:
        return VerificationResult.error_result("Pred: Failed to extract header/footer")

    differences = []

    def _compare_parts(part_name: str, gold_parts: list, pred_parts: list):
        # Key by (section, type): the extractor emits one part per section
        # per header/footer type. A type-only dict would collapse
        # multi-section documents onto the LAST section's part, silently
        # dropping every earlier section's evidence.
        def _by_key(parts):
            return {(part.get("section"), part.get("type")): part for part in parts}

        gold_by_key = _by_key(gold_parts)
        pred_by_key = _by_key(pred_parts)
        for part_key in sorted(
            set(gold_by_key) | set(pred_by_key),
            key=lambda k: (k[0] is None, k),
        ):
            gp = gold_by_key.get(part_key, {})
            pp = pred_by_key.get(part_key, {})

            for prop in [
                "text",
                "field_count",
                "page_field_count",
                "num_pages_field_count",
                "inline_shape_count",
                "shape_count",
            ]:
                gold_val = gp.get(prop)
                pred_val = pp.get(prop)
                if prop == "text":
                    gold_val = (gold_val or "").strip()
                    pred_val = (pred_val or "").strip()
                if gold_val != pred_val:
                    differences.append(
                        {
                            "part": part_name,
                            "section": part_key[0],
                            "type": part_key[1],
                            "property": prop,
                            "gold": gold_val,
                            "pred": pred_val,
                        }
                    )

    _compare_parts("header", gold_hf.get("headers", []), pred_hf.get("headers", []))
    _compare_parts("footer", gold_hf.get("footers", []), pred_hf.get("footers", []))

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Header/footer composition mismatch ({len(differences)} difference(s))",
            details={"differences": differences},
        )

    return VerificationResult(
        passed=True,
        message="PASS: Both header and footer content match",
        details={
            "header_count": len(gold_hf.get("headers", [])),
            "footer_count": len(gold_hf.get("footers", [])),
        },
    )
