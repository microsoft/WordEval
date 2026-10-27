"""preservation / structured with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import WORDPROCESSINGML_NS
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.preservation import visual as preservation_visual
from pathlib import Path
from typing import Any, Dict, Optional
import re
import xml.etree.ElementTree as ET
import zipfile


def saved_structured_target_present(docx_path: Path, kind: str) -> bool:
    """Return whether a saved package contains the structured target."""
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
    if kind == "chart":
        return any(name.startswith("word/charts/chart") and name.endswith(".xml") for name in names)
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    if kind == "equation":
        equation_tag = "{http://schemas.openxmlformats.org/officeDocument/2006/math}" "oMath"
        return any(node.tag == equation_tag for node in root.iter())
    if kind == "table":
        table_tag = f"{{{WORDPROCESSINGML_NS}}}tbl"
        return any(node.tag == table_tag for node in root.iter())
    raise ValueError(f"Unknown structured-object guardrail kind: {kind}")


def verify_structured_object_preservation(
    pred_path: Path,
    gold_path: Path,
    kind: str,
    *,
    allow_chart_terminal_empty: bool = False,
    read_page_info,
) -> VerificationResult:
    section_numeric_tolerance_twips = 20
    length_fields = {
        "page_size": {"w", "h"},
        "page_margins": {"top", "right", "bottom", "left", "header", "footer", "gutter"},
        "columns": {"space"},
        "document_grid": {"linePitch"},
    }

    def section_values_equivalent(gold_value: Any, pred_value: Any, path=()) -> bool:
        if isinstance(gold_value, dict) and isinstance(pred_value, dict):
            return gold_value.keys() == pred_value.keys() and all(
                section_values_equivalent(gold_value[key], pred_value[key], (*path, key))
                for key in gold_value
            )
        if isinstance(gold_value, list) and isinstance(pred_value, list):
            return len(gold_value) == len(pred_value) and all(
                section_values_equivalent(gold_item, pred_item, path)
                for gold_item, pred_item in zip(gold_value, pred_value)
            )
        if (
            isinstance(gold_value, str)
            and isinstance(pred_value, str)
            and re.fullmatch(r"-?\d+", gold_value)
            and re.fullmatch(r"-?\d+", pred_value)
        ):
            # Only physical lengths use twip rounding. Other numeric values
            # compare exactly while allowing equivalent integer spellings.
            tolerance = (
                section_numeric_tolerance_twips
                if len(path) == 2 and path[1] in length_fields.get(path[0], set())
                else 0
            )
            return abs(int(gold_value) - int(pred_value)) <= tolerance
        return gold_value == pred_value

    try:
        gold = preservation_visual.saved_visual_preservation_signature(
            gold_path, excluded_target_kind=kind
        )
        pred = preservation_visual.saved_visual_preservation_signature(
            pred_path, excluded_target_kind=kind
        )
    except Exception as exc:
        error = f"Structured-object preservation evidence unavailable: {exc}"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {error}",
            details={
                "error_owner": "evaluator",
                "guardrail_kind": kind,
                "match_mode": "saved_ooxml_structured_object_preservation",
            },
            error=error,
        )
    projection = preservation_visual.compare_saved_visual_preservation_signatures(
        gold,
        pred,
        allow_chart_terminal_empty=allow_chart_terminal_empty,
    )
    equivalent = projection["equivalent"] and section_values_equivalent(
        gold.get("section_layout"), pred.get("section_layout")
    )
    page_counts: Optional[Dict[str, int]] = None
    if equivalent and projection["requires_page_count_check"]:
        try:
            gold_page_info = read_page_info(str(gold_path))
            pred_page_info = read_page_info(str(pred_path))
            gold_page_count = gold_page_info.get("page_count")
            pred_page_count = pred_page_info.get("page_count")
        except Exception as exc:
            error = (
                "Structured-object paragraph-topology page-count evidence " f"unavailable: {exc}"
            )
            return VerificationResult(
                passed=False,
                message=f"ERROR: {error}",
                details={
                    "error_owner": "evaluator",
                    "guardrail_kind": kind,
                    "match_mode": ("saved_ooxml_structured_object_preservation"),
                    "projection_comparison": projection,
                },
                error=error,
            )
        if not isinstance(gold_page_count, int) or gold_page_count < 1:
            error = "Gold page-count evidence is unavailable"
            return VerificationResult(
                passed=False,
                message=f"ERROR: {error}",
                details={
                    "error_owner": "gold_fixture",
                    "guardrail_kind": kind,
                    "match_mode": ("saved_ooxml_structured_object_preservation"),
                    "projection_comparison": projection,
                },
                error=error,
            )
        page_counts = {
            "gold": gold_page_count,
            "pred": pred_page_count,
        }
        if pred_page_count != gold_page_count:
            equivalent = False
    if not equivalent:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Visible text/structure/page layout outside the "
                "requested structured object differs"
            ),
            details={
                "guardrail_kind": kind,
                "match_mode": "saved_ooxml_structured_object_preservation",
                "verification_source": "saved_ooxml_guardrail",
                "section_numeric_tolerance_twips": (section_numeric_tolerance_twips),
                "projection_comparison": projection,
                "page_counts": page_counts,
                "gold": gold,
                "pred": pred,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Structured-object collateral content is preserved",
        details={
            "guardrail_kind": kind,
            "match_mode": "saved_ooxml_structured_object_preservation",
            "verification_source": "saved_ooxml_guardrail",
            "section_numeric_tolerance_twips": (section_numeric_tolerance_twips),
            "projection_comparison": projection,
            "page_counts": page_counts,
        },
    )


def structured_gold_preflight(
    gold_path: Path, kind: Optional[str], *, expect_present: bool = True
) -> Optional[VerificationResult]:
    if not kind or not gold_path.exists():
        return None
    # Saved-fallback unit controls may intentionally contain only a chart
    # part.  Production packages have already passed Evaluator preflight;
    # defer package-level ownership/preservation for those minimal fixtures.
    with zipfile.ZipFile(gold_path) as package:
        if "word/document.xml" not in package.namelist():
            return None
    try:
        gold_has_target = saved_structured_target_present(gold_path, kind)
    except Exception as exc:
        error = f"Gold structured-object evidence could not be read: {exc}"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {error}",
            details={
                "error_owner": "gold_fixture",
                "guardrail_kind": kind,
                "missing_evidence": "saved_structured_object_target",
            },
            error=error,
        )
    if gold_has_target == expect_present:
        return None
    if expect_present:
        error = f"Gold has no {kind} target evidence for the requested " "object-effect route"
        missing_evidence = "saved_structured_object_target"
    else:
        # Destructive routes (e.g. TableDelete) must remove the target;
        # a gold that still contains it is a broken fixture.
        error = (
            f"Gold unexpectedly retains a {kind} target for the requested "
            "destructive object-effect route"
        )
        missing_evidence = "unexpected_structured_object_target"
    return VerificationResult(
        passed=False,
        message=f"ERROR: {error}",
        details={
            "error_owner": "gold_fixture",
            "guardrail_kind": kind,
            "missing_evidence": missing_evidence,
        },
        error=error,
    )


def finalize_structured_result(
    result: VerificationResult,
    pred_path: Path,
    gold_path: Path,
    kind: Optional[str],
    spec: Optional[Dict] = None,
    *,
    composition_steps,
    read_page_info,
) -> VerificationResult:
    if not kind or not result.passed or not pred_path.exists() or not gold_path.exists():
        return result
    with zipfile.ZipFile(pred_path) as pred_package, zipfile.ZipFile(gold_path) as gold_package:
        if (
            "word/document.xml" not in pred_package.namelist()
            or "word/document.xml" not in gold_package.namelist()
        ):
            return result
    preservation = verify_structured_object_preservation(
        pred_path,
        gold_path,
        kind,
        allow_chart_terminal_empty=(
            kind == "chart"
            and any(step.get("operation") == "insert_chart" for step in composition_steps(spec))
        ),
        read_page_info=read_page_info,
    )
    if not preservation.passed:
        return preservation
    result.details["structured_preservation"] = preservation.details
    return result
