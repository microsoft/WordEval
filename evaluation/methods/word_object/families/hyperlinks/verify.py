"""families / hyperlinks / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.hyperlinks import (
    extract_saved as families_hyperlinks_extract_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, Optional
import re
import xml.etree.ElementTree as ET
import zipfile
from urllib.parse import urlsplit, urlunsplit


def canonical_hyperlink_address(address):
    """RFC 3986 6.2.2.1: fold scheme/host, preserve path/query/userinfo case."""
    value = (address or '').strip()
    def percent(match):
        char = chr(int(match.group()[1:],16))
        return char if char.isascii() and (char.isalnum() or char in '-._~') else match.group().upper()
    value = re.sub(r'%[0-9A-Fa-f]{2}',percent,value)
    parts = urlsplit(value)
    if not parts.netloc:
        return value
    userinfo, separator, hostport = parts.netloc.rpartition('@')
    authority = (userinfo + separator if separator else '') + hostport.lower()
    path = parts.path
    if parts.scheme.lower() in {'http', 'https'} and path == '':
        path = '/'
    return urlunsplit((parts.scheme.lower(), authority, path, parts.query, parts.fragment))


def verify_hyperlink_saved_guardrail(pred_path: Path, gold_path: Path) -> VerificationResult:
    gold = families_hyperlinks_extract_saved.saved_hyperlink_guardrail_signature(gold_path)
    pred = families_hyperlinks_extract_saved.saved_hyperlink_guardrail_signature(pred_path)
    failures = []
    if pred["outside_hyperlink_formatting"] != gold["outside_hyperlink_formatting"]:
        failures.append("visible formatting outside hyperlink differs")
    if pred["hyperlink_extra_formatting"] != gold["hyperlink_extra_formatting"]:
        failures.append("unrequested hyperlink text formatting differs")
    if gold["standard_range_styles"]:
        if len(pred["standard_range_styles"]) != len(gold["standard_range_styles"]):
            failures.append("saved range-hyperlink count differs")
        elif any(
            gold_standard and not pred_standard
            for gold_standard, pred_standard in zip(
                gold["standard_range_styles"], pred["standard_range_styles"]
            )
        ):
            failures.append("standard blue hyperlink style is missing")
    if pred["pictures"] != gold["pictures"]:
        failures.append("linked picture content or visible layout differs")
    return VerificationResult(
        passed=not failures,
        message=(
            "PASS: Saved hyperlink formatting and linked content are preserved"
            if not failures
            else f"FAIL: Saved hyperlink guardrail differs ({len(failures)} check(s))"
        ),
        details={
            "match_mode": "saved_ooxml_hyperlink_guardrail",
            "failures": failures,
            "gold": gold,
            "pred": pred,
        },
    )


def verify_hyperlinks(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_hyperlink_info
) -> VerificationResult:
    """
    Verify hyperlink properties match between pred and gold.

    Handles both insert and delete/edit cases:
    - Insert: Checks hyperlink count, addresses, and display text match
    - Delete: Checks all hyperlinks removed, display text preserved
    - Edit: Checks hyperlink address unchanged, display text updated

    Scheme/host comparison is case-insensitive; destination paths and anchors are not.
    """
    if (spec or {}).get("idMso") == "HyperlinkRemove":
        return policies_content_operations.com_primary_with_saved_guardrail(
            lambda: verify_hyperlink_remove_com(
                pred_path, gold_path, read_hyperlink_info=read_hyperlink_info
            ),
            lambda: verify_hyperlink_remove_saved(pred_path, gold_path),
        )

    gold_info = read_hyperlink_info(str(gold_path))
    pred_info = read_hyperlink_info(str(pred_path))

    for owner, state in (("Gold", gold_info), ("Prediction", pred_info)):
        if (state.get('error') or state.get('hyperlink_count') is None
                or len(state.get('hyperlinks', [])) != state.get('hyperlink_count')):
            return policies_content_operations.error_with_details(
                f'{owner} hyperlink observation is incomplete', {'observation': state})

    gold_count = gold_info["hyperlink_count"]
    pred_count = pred_info["hyperlink_count"]

    # Compare hyperlink count
    if gold_count != pred_count:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Hyperlink count mismatch (pred={pred_count}, gold={gold_count})",
            details={
                "pred_count": pred_count,
                "gold_count": gold_count,
                "pred_hyperlinks": pred_info["hyperlinks"],
                "gold_hyperlinks": gold_info["hyperlinks"],
            },
        )

    # If gold has no hyperlinks (delete case), verify display text preserved
    if gold_count == 0:
        # Check that document text matches (display text kept)
        gold_text = gold_info.get("document_text", "").strip()
        pred_text = pred_info.get("document_text", "").strip()

        if gold_text and pred_text and gold_text != pred_text:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Document text differs after hyperlink removal",
                details={
                    "pred_text": pred_text,
                    "gold_text": gold_text,
                },
            )

        return VerificationResult(
            passed=True,
            message=f"PASS: All hyperlinks removed, display text preserved",
            details={
                "hyperlink_count": 0,
                "document_text": pred_text,
                "residual_hyperlink_style": bool(
                    pred_info.get("has_residual_hyperlink_style", False)
                ),
            },
        )

    # Compare each hyperlink
    for i, (gold_hl, pred_hl) in enumerate(zip(gold_info["hyperlinks"], pred_info["hyperlinks"])):
        gold_addr = canonical_hyperlink_address(gold_hl["address"])
        pred_addr = canonical_hyperlink_address(pred_hl["address"])

        if gold_addr != pred_addr:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Hyperlink {i+1} address mismatch",
                details={
                    "hyperlink_index": i + 1,
                    "pred_address": pred_hl["address"],
                    "gold_address": gold_hl["address"],
                },
            )

        if (gold_hl.get('sub_address') or '') != (pred_hl.get('sub_address') or ''):
            return VerificationResult(passed=False, message=f'FAIL: Hyperlink {i+1} anchor mismatch',
                details={'gold':gold_hl.get('sub_address'),'pred':pred_hl.get('sub_address')})

        # For shape hyperlinks (type=1, msoHyperlinkShape), skip text_to_display
        # comparison since images don't have meaningful display text
        if gold_hl["type"] != 1 and gold_hl["text_to_display"] != pred_hl["text_to_display"]:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Hyperlink {i+1} display text mismatch",
                details={
                    "hyperlink_index": i + 1,
                    "pred_display": pred_hl["text_to_display"],
                    "gold_display": gold_hl["text_to_display"],
                },
            )

        if gold_hl["type"] != pred_hl["type"]:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Hyperlink {i+1} type mismatch (pred={pred_hl['type']}, gold={gold_hl['type']})",
                details={
                    "hyperlink_index": i + 1,
                    "pred_type": pred_hl["type"],
                    "gold_type": gold_hl["type"],
                },
            )

        # Check character style (Hyperlink style gives blue color + underline)
        # Compare font color and underline to detect missing hyperlink formatting.
        # Gold generated via COM Hyperlinks.Add may not set explicit CharacterStyle,
        # but Word renders it with blue color and underline. Pred using python-docx
        # may create a functional hyperlink without the visual styling.
        gold_color = gold_hl.get("font_color")
        pred_color = pred_hl.get("font_color")
        gold_underline = gold_hl.get("font_underline")
        pred_underline = pred_hl.get("font_underline")

        # If pred has black text (no hyperlink styling) while gold has colored/undefined text
        # -16777216 = black (0xFF000000), 9999999 = wdUndefined (mixed formatting)
        # A properly styled hyperlink should NOT be plain black with no underline
        if pred_color is not None and pred_color == -16777216 and pred_underline == 0:
            # Pred has black text with no underline — missing hyperlink styling
            if gold_color != -16777216 or gold_underline != 0:
                return VerificationResult(
                    passed=False,
                    message=f"FAIL: Hyperlink {i+1} missing hyperlink styling (black text, no underline)",
                    details={
                        "hyperlink_index": i + 1,
                        "pred_font_color": pred_color,
                        "gold_font_color": gold_color,
                        "pred_font_underline": pred_underline,
                        "gold_font_underline": gold_underline,
                        "pred_style": pred_hl.get("character_style", ""),
                        "gold_style": gold_hl.get("character_style", ""),
                    },
                )

        # Also check explicit character style if both are set
        gold_style = gold_hl.get("character_style", "")
        pred_style = pred_hl.get("character_style", "")
        if gold_style and pred_style and gold_style != pred_style:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Hyperlink {i+1} character style mismatch (pred='{pred_style}', gold='{gold_style}')",
                details={
                    "hyperlink_index": i + 1,
                    "pred_style": pred_style,
                    "gold_style": gold_style,
                    "pred_font_color": pred_color,
                    "gold_font_color": gold_color,
                },
            )

    # Compare document text to catch surrounding text differences
    gold_text = gold_info.get("document_text", "").strip()
    pred_text = pred_info.get("document_text", "").strip()

    if gold_text != pred_text:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Document text differs",
            details={
                "pred_text": pred_text,
                "gold_text": gold_text,
                "hyperlink_count": gold_count,
            },
        )

    saved_guardrail = None
    if (spec or {}).get("idMso") in {"HyperlinkInsert", "HyperlinkEdit"}:
        saved_guardrail = verify_hyperlink_saved_guardrail(pred_path, gold_path)
        if not saved_guardrail.passed:
            return saved_guardrail

    # Build summary message
    if gold_count == 1:
        hl = gold_info["hyperlinks"][0]
        msg = f"PASS: Hyperlink matches (address='{hl['address']}', display='{hl['text_to_display']}')"
    else:
        msg = f"PASS: All {gold_count} hyperlinks match"

    return VerificationResult(
        passed=True,
        message=msg,
        details={
            "hyperlink_count": gold_count,
            "hyperlinks": pred_info["hyperlinks"],
            "saved_guardrail": (saved_guardrail.details if saved_guardrail is not None else None),
        },
    )


def verify_hyperlink_remove_com(
    pred_path: Path, gold_path: Path, *, read_hyperlink_info
) -> VerificationResult:
    gold = read_hyperlink_info(str(gold_path))
    pred = read_hyperlink_info(str(pred_path))
    failures = []
    if pred.get("hyperlink_count") != gold.get("hyperlink_count"):
        failures.append("hyperlink count differs")
    if pred.get("document_text") != gold.get("document_text"):
        failures.append("display text differs")
    return VerificationResult(
        passed=not failures,
        message=(
            "PASS: All Word Object Model hyperlinks removed, display text preserved"
            if not failures
            else f"FAIL: Word Object Model hyperlink removal differs ({len(failures)} check(s))"
        ),
        details={"failures": failures, "gold": gold, "pred": pred},
    )


def verify_hyperlink_remove_saved(pred_path: Path, gold_path: Path) -> VerificationResult:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def saved_signature(path: Path) -> Dict[str, Any]:
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
        hyperlink_elements = list(root.iter(f"{{{w}}}hyperlink"))
        field_count = sum(
            1
            for node in root.iter(f"{{{w}}}instrText")
            if re.search(r"\bHYPERLINK\b", node.text or "", re.IGNORECASE)
        )
        return {
            "hyperlink_count": len(hyperlink_elements) + field_count,
            "relationship_ids": sorted(
                node.get(f"{{{r}}}id") for node in hyperlink_elements if node.get(f"{{{r}}}id")
            ),
            "document_text": "".join(
                node.text or "" for node in root.findall(".//w:body//w:t", {"w": w})
            ),
            "residual_hyperlink_style": any(
                node.get(f"{{{w}}}val") == "Hyperlink" for node in root.iter(f"{{{w}}}rStyle")
            ),
        }

    gold = saved_signature(gold_path)
    pred = saved_signature(pred_path)
    failures = []
    if pred["hyperlink_count"] != gold["hyperlink_count"]:
        failures.append("hyperlink count differs")
    if pred["document_text"] != gold["document_text"]:
        failures.append("display text differs")
    return VerificationResult(
        passed=not failures,
        message=(
            "PASS: Saved hyperlink removal matches"
            if not failures
            else f"FAIL: Saved hyperlink removal differs ({len(failures)} check(s))"
        ),
        details={
            "match_mode": "saved_ooxml_hyperlink_removal",
            "failures": failures,
            "gold": gold,
            "pred": pred,
        },
    )
