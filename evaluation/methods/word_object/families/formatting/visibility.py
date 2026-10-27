"""families / formatting / visibility with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Any, Dict, List, Optional


def composition_requires_target_visibility(spec: Dict) -> bool:
    """Return whether a composition replaces target glyph paint.

    Gradient-fill compositions need an explicit visibility invariant:
    matching the fill token and angle does not prove that Word paints any
    glyph.  This semantic condition is operation-scoped rather than tied
    to a case, model, path, or saved-file fingerprint.
    """
    return any(
        intent.get("operation") == "fill_gradient"
        for intent in spec.get("composition", {}).get("atomic_intents", [])
    )


def glyph_paint_contribution(state: Dict[str, Any], prefix: str) -> Optional[bool]:
    visible = state.get(f"{prefix}_visible")
    transparency = state.get(f"{prefix}_transparency")
    if visible is False:
        return False
    if visible is not True or transparency is None:
        return None
    if transparency >= 0.999:
        return False
    if prefix == "outline":
        weight = state.get("outline_weight")
        if weight is not None and weight <= 0:
            return False
    return True


def effective_target_glyph_visible(state: Dict[str, Any]) -> Optional[bool]:
    if not state.get("found"):
        return False
    if state.get("hidden") is True:
        return False
    contributions = [
        glyph_paint_contribution(state, "fill"),
        glyph_paint_contribution(state, "outline"),
    ]
    if any(value is True for value in contributions):
        return True
    if all(value is False for value in contributions):
        return False
    if (
        state.get("fill_type") == 1
        and state.get("fill_transparency") is None
        and state.get("font_color") is not None
    ):
        # Word's ordinary solid text can expose Font.Color while the
        # Office Fill.Transparency member is unavailable.  This still
        # proves glyph paint exists; do not use the color value as a
        # contrast judgment.
        return True
    return None


def verify_target_glyph_visibility(
    pred_path: Path, gold_path: Path, spec: Dict, target_text: str, *, read_target_glyph_state
) -> Optional[VerificationResult]:
    """Require determinate Word-effective glyph paint on the target.

    This is not a contrast score.  Any color/theme/gradient endpoints pass
    when Word reports non-transparent fill or outline paint.  A failure is
    limited to a missing/hidden target or the determinate state where both
    glyph paint channels are absent.  Incomplete COM evidence remains an
    evaluator error instead of being guessed into a model failure.
    """
    if not composition_requires_target_visibility(spec):
        return None

    gold_state = read_target_glyph_state(str(gold_path), target_text)
    pred_state = read_target_glyph_state(str(pred_path), target_text)
    gold_visible = effective_target_glyph_visible(gold_state)
    pred_visible = effective_target_glyph_visible(pred_state)
    details = {
        "match_mode": "word_com_target_glyph_visibility",
        "target_text": target_text,
        "gold_visible": gold_visible,
        "pred_visible": pred_visible,
        "gold_state": gold_state,
        "pred_state": pred_state,
    }

    if gold_visible is not True:
        message = (
            "Gold target has no effective glyph rendering"
            if gold_visible is False
            else "Gold target glyph visibility is indeterminate"
        )
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details=details,
            error=message,
        )
    if pred_visible is False:
        return VerificationResult(
            passed=False,
            message=(f'FAIL: Target text "{target_text}" has no effective ' "glyph rendering"),
            details=details,
        )
    if pred_visible is None:
        message = "Prediction target glyph visibility is indeterminate"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details=details,
            error=message,
        )
    return VerificationResult(
        passed=True,
        message=f'PASS: Target text "{target_text}" has effective glyph rendering',
        details=details,
    )


def verify_control_glyph_visibility(
    pred_path: Path,
    gold_path: Path,
    spec: Dict,
    control_texts: List[str],
    *,
    read_target_glyph_state,
) -> Optional[VerificationResult]:
    """Preserve determinate glyph rendering for non-target control text."""
    if not composition_requires_target_visibility(spec):
        return None

    rows = []
    for text in control_texts:
        gold_state = read_target_glyph_state(str(gold_path), text)
        pred_state = read_target_glyph_state(str(pred_path), text)
        gold_visible = effective_target_glyph_visible(gold_state)
        pred_visible = effective_target_glyph_visible(pred_state)
        rows.append(
            {
                "text": text,
                "gold_visible": gold_visible,
                "pred_visible": pred_visible,
                "gold_state": gold_state,
                "pred_state": pred_state,
            }
        )

    details = {
        "match_mode": "word_com_control_glyph_visibility",
        "controls": rows,
    }
    indeterminate_gold = [row["text"] for row in rows if row["gold_visible"] is not True]
    if indeterminate_gold:
        message = "Gold control glyph visibility is missing or indeterminate: " + ", ".join(
            indeterminate_gold
        )
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details=details,
            error=message,
        )
    missing_pred = [row["text"] for row in rows if row["pred_visible"] is False]
    if missing_pred:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Non-target control text has no effective glyph "
                "rendering: " + ", ".join(missing_pred)
            ),
            details=details,
        )
    indeterminate_pred = [row["text"] for row in rows if row["pred_visible"] is None]
    if indeterminate_pred:
        message = "Prediction control glyph visibility is indeterminate: " + ", ".join(
            indeterminate_pred
        )
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details=details,
            error=message,
        )
    return VerificationResult(
        passed=True,
        message="PASS: Non-target control text retains glyph rendering",
        details=details,
    )
