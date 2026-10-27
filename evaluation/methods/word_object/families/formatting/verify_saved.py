"""families / formatting / verify_saved with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.formatting import (
    compare as families_formatting_compare,
)
from evaluation.methods.word_object.families.formatting import (
    extract_saved as families_formatting_extract_saved,
)
from pathlib import Path
from typing import Dict


def verify_ligatures(pred_path: Path, gold_path: Path, spec: Dict) -> VerificationResult:
    """Verify saved OpenType ligature settings without opening Word."""
    gold_fonts = families_formatting_extract_saved.extract_saved_ligature_signatures(gold_path)
    pred_fonts = families_formatting_extract_saved.extract_saved_ligature_signatures(pred_path)

    gold_has_ligature_evidence = bool(gold_fonts) and any(
        any(value is not None for value in paragraph.get("ligatures", []))
        for paragraph in gold_fonts
    )
    if not gold_has_ligature_evidence:
        return VerificationResult.error_result("Gold saved ligature evidence is missing")

    differences = []
    if gold_fonts != pred_fonts:
        differences.append(
            {
                "property": "saved_ligature_signatures",
                "gold": gold_fonts,
                "pred": pred_fonts,
            }
        )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} ligature difference(s)",
            details={
                "differences": differences,
                "gold": gold_fonts,
                "pred": pred_fonts,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Ligature settings match",
        details={
            "paragraph_count": len(gold_fonts),
            "match_mode": "saved_ooxml_ligatures",
        },
    )


def verify_saved_font_composition(
    pred_path: Path, gold_path: Path, spec: Dict
) -> VerificationResult:
    target = spec.get("precondition", {}).get("select_text", "Target")
    gold = families_formatting_extract_saved.extract_saved_target_run(gold_path, target)
    pred = families_formatting_extract_saved.extract_saved_target_run(pred_path, target)
    if not gold["target_runs"]:
        return VerificationResult.error_result(
            f'Gold saved target-run evidence is missing for "{target}"'
        )
    differences = []
    if gold["full_text"] != pred["full_text"]:
        differences.append(
            {
                "property": "document_text",
                "gold": gold["full_text"],
                "pred": pred["full_text"],
            }
        )
    if gold["target_runs"] != pred["target_runs"]:
        differences.append(
            {
                "property": "target_run_formatting",
                "gold": gold["target_runs"],
                "pred": pred["target_runs"],
            }
        )
    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Saved font composition differs ({len(differences)} check(s))",
            details={
                "differences": differences,
                "match_mode": "saved_ooxml_target_run",
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Saved target-run font composition matches",
        details={"match_mode": "saved_ooxml_target_run", "signature": gold},
    )


def verify_saved_target_text_effect(
    pred_path: Path, gold_path: Path, spec: Dict
) -> VerificationResult:
    target = spec.get("precondition", {}).get("select_text", "Target")
    idMso = spec.get("idMso", "")
    gold = families_formatting_extract_saved.extract_saved_target_text_effect(
        gold_path, target, idMso
    )
    pred = families_formatting_extract_saved.extract_saved_target_text_effect(
        pred_path, target, idMso
    )
    gold_runs = gold.get("target_runs", [])
    if idMso == "TextHighlightColorPicker":
        gold_has_effect = any(run.get("effect") is not None for run in gold_runs)
    elif idMso == "TextOutlineDashesGallery":
        gold_has_effect = any(
            (run.get("effect") or {}).get("outline_visible")
            and (run.get("effect") or {}).get("dash") is not None
            for run in gold_runs
        )
    elif idMso == "TextFillGradientGallery":
        gold_has_effect = any(
            (run.get("effect") or {}).get("gradient_present")
            and (run.get("effect") or {}).get("stops")
            and (run.get("effect") or {}).get("angle") is not None
            for run in gold_runs
        )
    else:
        gold_has_effect = bool(gold_runs)
    if not gold_has_effect:
        return VerificationResult.error_result(
            f'Gold saved {idMso} target evidence is missing for "{target}"'
        )
    if families_formatting_compare.round_dot_text_outline_requested(spec):
        for extracted in (gold, pred):
            for run in extracted.get("target_runs", []):
                effect = run.get("effect") or {}
                if effect.get("dash") in {"dot", "sysDot"}:
                    effect["dash"] = "round_dot_semantic"
    if gold != pred:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Saved {idMso} signature differs",
            details={
                "match_mode": "saved_ooxml_text_effect",
                "gold": gold,
                "pred": pred,
            },
        )
    return VerificationResult(
        passed=True,
        message=f"PASS: Saved {idMso} signature matches",
        details={
            "match_mode": "saved_ooxml_text_effect",
            "signature": gold,
        },
    )


def verify_saved_multilevel_list(pred_path: Path, gold_path: Path) -> VerificationResult:
    gold = families_formatting_extract_saved.extract_saved_multilevel_list(gold_path)
    pred = families_formatting_extract_saved.extract_saved_multilevel_list(pred_path)
    if gold.get("invalid") or not gold.get("paragraphs") or not gold.get("levels"):
        return VerificationResult.error_result(
            "Gold saved multilevel-list evidence is missing or invalid"
        )
    if gold != pred:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved multilevel-list definition differs",
            details={
                "match_mode": "saved_ooxml_numbering",
                "gold": gold,
                "pred": pred,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Saved multilevel-list definition matches",
        details={"match_mode": "saved_ooxml_numbering", "signature": gold},
    )


def verify_quick_styles_saved(pred_path: Path, gold_path: Path, spec: Dict) -> VerificationResult:
    """Verify the style token and absence of conflicting direct overrides."""
    gold_signatures = families_formatting_extract_saved.extract_quick_style_signatures(gold_path)
    pred_signatures = families_formatting_extract_saved.extract_quick_style_signatures(pred_path)

    if not gold_signatures or not any(row.get("style_id") for row in gold_signatures):
        return VerificationResult.error_result("Gold saved Quick Styles evidence is missing")

    if len(gold_signatures) != len(pred_signatures):
        return VerificationResult(
            passed=False,
            message=(
                f"FAIL: Paragraph count mismatch "
                f"(gold: {len(gold_signatures)}, pred: {len(pred_signatures)})"
            ),
            details={"gold": gold_signatures, "pred": pred_signatures},
        )

    differences = []
    for i, (gold, pred) in enumerate(zip(gold_signatures, pred_signatures)):
        if gold["text"] != pred["text"]:
            differences.append(
                {
                    "property": "paragraph_text",
                    "gold": gold["text"],
                    "pred": pred["text"],
                    "location": f"paragraph {i + 1}",
                }
            )
        if gold["style_id"] != pred["style_id"]:
            differences.append(
                {
                    "property": "paragraph_style",
                    "gold": gold["style_id"],
                    "pred": pred["style_id"],
                    "location": f"paragraph {i + 1}",
                }
            )
        for prop in ("paragraph_direct_formatting", "run_direct_formatting"):
            if gold[prop] != pred[prop]:
                differences.append(
                    {
                        "property": prop,
                        "gold": gold[prop],
                        "pred": pred[prop],
                        "location": f"paragraph {i + 1}",
                    }
                )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} paragraph style difference(s)",
            details={
                "differences": differences,
                "gold": gold_signatures,
                "pred": pred_signatures,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Paragraph styles and direct formatting match",
        details={"paragraph_signatures": gold_signatures},
    )
