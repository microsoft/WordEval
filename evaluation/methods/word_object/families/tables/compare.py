"""families / tables / compare with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from typing import Any, Dict, List, Optional, Tuple


def preferred_column_spread(table: Dict[str, Any]) -> Optional[float]:
    representatives = []
    width_type = None
    for column in table.get("preferred_columns", []):
        valid = [(kind, value) for kind, value in column if value is not None]
        if not valid:
            return None
        kinds = {kind for kind, _value in valid}
        if len(kinds) != 1:
            return None
        kind = valid[0][0]
        if width_type is None:
            width_type = kind
        elif kind != width_type:
            return None
        representatives.append(sum(value for _kind, value in valid) / len(valid))
    if not representatives:
        return None
    return max(representatives) - min(representatives)


def rendered_table_width_tolerances(
    gold_widths: List[float], pred_widths: List[float]
) -> Dict[str, float]:
    """Return Word-layout tolerances calibrated as scale-aware bounds.

    The retained correct reopen drift is 2.5 pt on a roughly 468 pt
    window-width table (0.535%).  A 0.6% total-width bound clears that
    serialization/layout rounding while staying far below the observed
    fixed/contents and visibly drifting controls.  Equal-column spread is
    independently bounded to 0.3% of one rendered column (at least 0.5 pt).
    """
    gold_total = sum(gold_widths)
    pred_total = sum(pred_widths)
    pred_average = pred_total / len(pred_widths) if pred_widths else 0.0
    return {
        "total_width": max(1.0, 0.006 * max(gold_total, pred_total)),
        "column_spread": max(0.5, 0.003 * pred_average),
    }


def verify_table_column_insert_semantics(
    gold_tables: List[Dict], pred_tables: List[Dict], idMso: str
) -> VerificationResult:
    """Verify inserted-column outcomes without unrelated cell metadata."""
    differences = []
    if len(gold_tables) != len(pred_tables):
        differences.append(
            {
                "property": "table_count",
                "gold": len(gold_tables),
                "pred": len(pred_tables),
            }
        )
    for index, (gold, pred) in enumerate(zip(gold_tables, pred_tables)):
        for prop in ("row_count", "column_count", "cell_contents"):
            if gold.get(prop) != pred.get(prop):
                differences.append(
                    {
                        "table_index": index,
                        "property": prop,
                        "gold": gold.get(prop),
                        "pred": pred.get(prop),
                    }
                )
        gold_widths = gold.get("column_widths") or []
        pred_widths = pred.get("column_widths") or []
        if len(gold_widths) != len(pred_widths) or any(
            abs(gold_value - pred_value) > 2.0
            for gold_value, pred_value in zip(gold_widths, pred_widths)
        ):
            differences.append(
                {
                    "table_index": index,
                    "property": "column_widths",
                    "gold": gold_widths,
                    "pred": pred_widths,
                }
            )
    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Inserted-column semantics differ ({len(differences)} difference(s))",
            details={"idMso": idMso, "differences": differences},
        )
    return VerificationResult(
        passed=True,
        message="PASS: Inserted column structure, content, and widths match",
        details={"idMso": idMso, "table_count": len(pred_tables)},
    )


def verify_insert_then_border_semantics(
    gold_tables: List[Dict], pred_tables: List[Dict]
) -> VerificationResult:
    """Verify table insertion plus outer border style only.

    The instruction does not specify border color or weight. Those are
    theme/default choices and must not make a correct double outer border
    fail.
    """
    differences = []
    if len(gold_tables) != len(pred_tables):
        differences.append(
            {
                "property": "table_count",
                "gold": len(gold_tables),
                "pred": len(pred_tables),
            }
        )
    for index, (gold, pred) in enumerate(zip(gold_tables, pred_tables)):
        for prop in ("row_count", "column_count", "cell_contents"):
            if gold.get(prop) != pred.get(prop):
                differences.append(
                    {
                        "table_index": index,
                        "property": prop,
                        "gold": gold.get(prop),
                        "pred": pred.get(prop),
                    }
                )
        for border_name in ("top", "left", "bottom", "right"):
            gold_style = (gold.get("borders") or {}).get(border_name, {}).get("line_style")
            pred_style = (pred.get("borders") or {}).get(border_name, {}).get("line_style")
            if gold_style != pred_style:
                differences.append(
                    {
                        "table_index": index,
                        "property": f"border_{border_name}_line_style",
                        "gold": gold_style,
                        "pred": pred_style,
                    }
                )
    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Inserted table border semantics differ ({len(differences)} difference(s))",
            details={"differences": differences},
        )
    return VerificationResult(
        passed=True,
        message="PASS: Inserted table and outer border style match",
        details={"table_count": len(pred_tables)},
    )


def verify_additive_table_style_direction_border(
    gold_tables: List[Dict], pred_tables: List[Dict]
) -> VerificationResult:
    """Verify only the requested style, direction, and outer borders."""
    properties = (
        "row_count",
        "column_count",
        "cell_contents",
        "style",
        "cell_text_orientations",
    )
    differences = []
    if len(gold_tables) != len(pred_tables):
        differences.append(
            {
                "property": "table_count",
                "gold": len(gold_tables),
                "pred": len(pred_tables),
            }
        )
    for index, (gold, pred) in enumerate(zip(gold_tables, pred_tables)):
        for prop in properties:
            if gold.get(prop) != pred.get(prop):
                differences.append(
                    {
                        "table_index": index,
                        "property": prop,
                        "gold": gold.get(prop),
                        "pred": pred.get(prop),
                    }
                )
        for name in ("top", "left", "bottom", "right"):
            gold_style = (gold.get("borders") or {}).get(name, {}).get("line_style")
            pred_style = (pred.get("borders") or {}).get(name, {}).get("line_style")
            if gold_style != pred_style:
                differences.append(
                    {
                        "table_index": index,
                        "property": f"border_{name}_line_style",
                        "gold": gold_style,
                        "pred": pred_style,
                    }
                )
    if differences:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Requested table style/direction/border semantics "
                f"differ ({len(differences)} difference(s))"
            ),
            details={"differences": differences},
        )
    return VerificationResult(
        passed=True,
        message="PASS: Requested table style, direction, and outer border match",
        details={"table_count": len(pred_tables)},
    )


def verify_inserted_table_properties(
    gold_tables: List[Dict],
    pred_tables: List[Dict],
    properties: Tuple[str, ...],
    operation_label: str,
) -> VerificationResult:
    """Compare only structure/content and the requested table operation.

    A table inserted through Word's UI can inherit a different default
    style and border set than a table inserted through COM. Those defaults
    are outside the contract for alignment-only compositions.
    """
    differences = []
    if len(gold_tables) != len(pred_tables):
        differences.append(
            {
                "property": "table_count",
                "gold": len(gold_tables),
                "pred": len(pred_tables),
            }
        )
    for index, (gold, pred) in enumerate(zip(gold_tables, pred_tables)):
        for prop in properties:
            if gold.get(prop) != pred.get(prop):
                differences.append(
                    {
                        "table_index": index,
                        "property": prop,
                        "gold": gold.get(prop),
                        "pred": pred.get(prop),
                    }
                )
    if differences:
        return VerificationResult(
            passed=False,
            message=(
                f"FAIL: Inserted table {operation_label} differs "
                f"({len(differences)} difference(s))"
            ),
            details={"differences": differences},
        )
    return VerificationResult(
        passed=True,
        message=f"PASS: Inserted table {operation_label} matches",
        details={"table_count": len(pred_tables)},
    )
