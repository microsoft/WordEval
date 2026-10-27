"""families / metadata / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re

DOCINFO_COMMANDS = {"DocInfoFileName", "DocInfoAuthor", "DocInfoDocTitle"}


def verify_docinfo(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_docinfo_state
) -> VerificationResult:
    """Verify real, correctly located DocInfo semantics in the header."""
    id_mso = spec.get("idMso", "") if spec else ""
    if id_mso not in DOCINFO_COMMANDS:
        return VerificationResult(
            passed=False, message=f"FAIL: Unsupported DocInfo command: {id_mso}", details={}
        )
    gold = read_docinfo_state(str(gold_path))
    pred = read_docinfo_state(str(pred_path))

    def preservation_differences() -> Dict[str, Any]:
        differences = {}
        for key in ("section_count", "body_text", "footers"):
            if pred.get(key) != gold.get(key):
                differences[key] = {"gold": gold.get(key), "pred": pred.get(key)}
        return differences

    if id_mso == "DocInfoFileName":

        def filename_fields(state: Dict[str, Any]) -> List[Dict[str, Any]]:
            return [
                {
                    **field,
                    "section_index": header.get("section_index"),
                    "header_type": header.get("type"),
                }
                for header in state.get("headers", [])
                for field in header.get("fields", [])
                if field.get("type") == 29
                or re.search(r"^\s*FILENAME\b", field.get("code") or "", re.I)
            ]

        gold_fields = filename_fields(gold)
        pred_fields = filename_fields(pred)
        if len(gold_fields) != 1:
            return policies_content_operations.error_with_details(
                "Gold must contain exactly one header FILENAME field",
                {"gold": gold, "gold_filename_fields": gold_fields},
            )
        if len(pred_fields) != 1:
            return VerificationResult(
                passed=False,
                message=(
                    "FAIL: Pred must contain exactly one header FILENAME field "
                    f"(found {len(pred_fields)})"
                ),
                details={"gold_fields": gold_fields, "pred_fields": pred_fields},
            )
        field = pred_fields[0]
        if re.search(r"\\p\b", field.get("code") or "", re.I):
            return VerificationResult(
                passed=False,
                message="FAIL: Pred FILENAME field uses the \\p full-path switch",
                details={"pred_field": field},
            )
        location = (field.get("section_index"), field.get("header_type"))
        gold_location = (
            gold_fields[0].get("section_index"),
            gold_fields[0].get("header_type"),
        )
        if location != gold_location:
            return VerificationResult(
                passed=False,
                message="FAIL: FILENAME field is in the wrong header story",
                details={"gold_location": gold_location, "pred_location": location},
            )
        refreshed = read_docinfo_state(str(pred_path), update_fields=True)
        refreshed_fields = filename_fields(refreshed)
        refreshed_result = (
            refreshed_fields[0].get("result", "") if len(refreshed_fields) == 1 else ""
        )
        if not refreshed_result or re.search(r"[\\/]", refreshed_result):
            return VerificationResult(
                passed=False,
                message="FAIL: Refreshed FILENAME does not render a current file name",
                details={
                    "submitted": field,
                    "refreshed_fields": refreshed_fields,
                },
            )
        collateral = preservation_differences()
        if collateral:
            return VerificationResult(
                passed=False,
                message="FAIL: DocInfo FileName insertion changed unrelated document state",
                details={"differences": collateral},
            )
        return VerificationResult(
            passed=True,
            message="PASS: Unique header FILENAME field has current-file semantics",
            details={
                "submitted_field": field,
                "refreshed_result": refreshed_result,
                "match_mode": "docinfo_filename_com_submitted_and_refreshed",
            },
        )

    expected_key = "Author" if id_mso == "DocInfoAuthor" else "Title"
    expected_xpath = "creator" if expected_key == "Author" else "title"

    def matching_controls(state: Dict[str, Any]) -> List[Dict[str, Any]]:
        matches = []
        for header in state.get("headers", []):
            for control in header.get("controls", []):
                title_or_tag = (control.get("title") or control.get("tag") or "").casefold()
                xpath = (control.get("xpath") or "").casefold()
                if title_or_tag == expected_key.casefold() or re.search(
                    rf"(?:^|[:/]){expected_xpath}$", xpath
                ):
                    matches.append(
                        {
                            **control,
                            "section_index": header.get("section_index"),
                            "header_type": header.get("type"),
                        }
                    )
        return matches

    gold_controls = matching_controls(gold)
    pred_controls = matching_controls(pred)
    valid_gold = [
        control
        for control in gold_controls
        if control.get("is_mapped")
        and control.get("source_exists")
        and expected_xpath in (control.get("xpath") or "").casefold()
    ]
    if len(valid_gold) != 1:
        return policies_content_operations.error_with_details(
            f"Gold lacks exactly one valid mapped {expected_key} control",
            {"gold_controls": gold_controls, "gold": gold},
        )
    valid_pred = [
        control
        for control in pred_controls
        if control.get("is_mapped")
        and control.get("source_exists")
        and expected_xpath in (control.get("xpath") or "").casefold()
    ]
    if len(valid_pred) != 1 or len(pred_controls) != 1:
        return VerificationResult(
            passed=False,
            message=(
                f"FAIL: Pred must contain exactly one correctly mapped "
                f"{expected_key} header control"
            ),
            details={
                "gold_controls": gold_controls,
                "pred_controls": pred_controls,
                "valid_pred_controls": valid_pred,
            },
        )
    gold_control = valid_gold[0]
    pred_control = valid_pred[0]
    if (pred_control.get("section_index"), pred_control.get("header_type")) != (
        gold_control.get("section_index"),
        gold_control.get("header_type"),
    ):
        return VerificationResult(
            passed=False,
            message=f"FAIL: {expected_key} control is in the wrong header story",
            details={"gold": gold_control, "pred": pred_control},
        )
    if not " ".join((pred_control.get("text") or "").split()):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Mapped {expected_key} control has no displayed value",
            details={"gold": gold_control, "pred": pred_control},
        )
    collateral = preservation_differences()
    if collateral:
        return VerificationResult(
            passed=False,
            message=f"FAIL: DocInfo {expected_key} insertion changed unrelated state",
            details={"differences": collateral},
        )
    return VerificationResult(
        passed=True,
        message=f"PASS: Unique mapped {expected_key} header control verified",
        details={
            "control": pred_control,
            "match_mode": "docinfo_xml_mapping_and_header_story",
        },
    )
