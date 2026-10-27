"""families / formatting / operations with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Dict, Optional
from ...observations import require_observations


def verify_paragraph_properties_composition(
    pred_path: Path,
    gold_path: Path,
    *,
    error_with_details,
    read_all_paragraphs_formatting,
    verify_dropcap,
) -> VerificationResult:
    """Verify paragraph spacing and drop cap match gold."""
    # Verify paragraph spacing
    spacing_result = verify_paragraph_spacing(
        pred_path,
        gold_path,
        spec=None,
        error_with_details=error_with_details,
        read_all_paragraphs_formatting=read_all_paragraphs_formatting,
    )
    if not spacing_result.passed:
        return spacing_result

    # Verify drop cap
    dropcap_result = verify_dropcap(pred_path, gold_path)
    if not dropcap_result.passed:
        return dropcap_result

    return VerificationResult(
        passed=True,
        message="PASS: Paragraph properties composition (spacing + drop cap)",
        details={
            "spacing": spacing_result.details,
            "dropcap": dropcap_result.details,
        },
    )


def verify_paragraph_spacing(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    error_with_details,
    read_all_paragraphs_formatting,
) -> VerificationResult:
    """
    Verify paragraph spacing matches gold for the target property only.

    Only checks the property specified in spec params (SpaceBefore or SpaceAfter),
    not both. This avoids false failures when the model correctly sets the target
    property but doesn't touch the other.
    """
    gold_paras = read_all_paragraphs_formatting(str(gold_path))
    pred_paras = read_all_paragraphs_formatting(str(pred_path))

    if gold_paras is None or pred_paras is None:
        return error_with_details(
            "Could not extract paragraph formatting",
            {
                "gold_extracted": gold_paras is not None,
                "pred_extracted": pred_paras is not None,
            },
        )

    if not gold_paras:
        return VerificationResult.error_result("No gold paragraphs available for spacing verification")

    if len(gold_paras) != len(pred_paras):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Paragraph count mismatch (pred={len(pred_paras)}, gold={len(gold_paras)})",
            details={"pred_count": len(pred_paras), "gold_count": len(gold_paras)},
        )

    # Determine which property to check from spec params
    target_prop = None
    if spec:
        verification = spec.get("verification", {})
        primary = verification.get("primary", {})
        params = primary.get("params", {})
        property_name = params.get("property_name", "")
        # Map COM property name to extraction key
        prop_map = {"SpaceBefore": "space_before", "SpaceAfter": "space_after"}
        target_prop = prop_map.get(property_name)

    # If no specific property, check both
    props_to_check = [target_prop] if target_prop else ["space_before", "space_after"]

    differences = []
    for i, (gp, pp) in enumerate(zip(gold_paras, pred_paras)):
        error = require_observations(
            {"gold": gp, "pred": pp}, props_to_check, context=f"paragraph {i + 1} spacing",
        )
        if error is not None:
            return error
        para_diffs = {}
        for prop in props_to_check:
            gold_val = gp.get(prop)
            pred_val = pp.get(prop)
            if gold_val is not None and pred_val is not None:
                # Allow 0.5pt tolerance for floating point
                if abs(gold_val - pred_val) > 0.5:
                    para_diffs[prop] = {"gold": gold_val, "pred": pred_val}
        if para_diffs:
            differences.append({"paragraph": i + 1, "differences": para_diffs})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Paragraph spacing mismatch in {len(differences)} paragraph(s)",
            details={"differences": differences},
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Paragraph spacing matches for all {len(gold_paras)} paragraphs",
        details={"paragraph_count": len(gold_paras)},
    )
