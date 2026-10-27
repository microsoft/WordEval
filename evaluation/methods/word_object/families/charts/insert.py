"""Chart insertion orchestration with explicit shared-policy dependencies."""

from pathlib import Path
from typing import Any, Callable, Dict, Optional

from ....base import VerificationResult
from .compare import expected_insert_contract, com_insert_failures, saved_insert_failures
from .extract_saved import extract_insertion_signature


def verify_saved(
    pred_path: Path,
    gold_path: Path,
    spec: Dict[str, Any],
    *,
    verify_context: Callable,
    error_result: Callable,
) -> VerificationResult:
    expected = expected_insert_contract(spec)
    try:
        gold = extract_insertion_signature(gold_path)
    except Exception as exc:
        return error_result(f"Cannot extract gold saved chart evidence: {exc}")
    gold_failures = saved_insert_failures(gold, expected)
    if gold_failures:
        return error_result(
            "Gold saved chart evidence is incomplete or invalid",
            {"failures": gold_failures, "gold": gold, "expected": expected},
        )
    try:
        pred = extract_insertion_signature(pred_path)
    except Exception as exc:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Predicted saved chart package is unreadable ({exc})",
            details={"gold": gold, "expected": expected},
        )
    failures = saved_insert_failures(pred, expected)
    context_failure = verify_context(pred_path, gold_path)
    if context_failure is not None:
        failures.append("surrounding document content/layout changed")
    return VerificationResult(
        passed=not failures,
        message=(
            "PASS: Saved chart graph, cache, workbook, and context are valid"
            if not failures
            else f"FAIL: Saved chart insertion differs ({'; '.join(failures)})"
        ),
        details={
            "match_mode": "saved_ooxml_chart_graph_cache_workbook",
            "failures": failures,
            "gold": gold,
            "pred": pred,
            "expected": expected,
            "context_failure": (context_failure.to_dict() if context_failure else None),
        },
    )


def verify_insert(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_com: Callable[[str], Dict],
    verify_saved: Callable,
    combine_evidence: Callable,
    error_result: Callable,
) -> VerificationResult:
    """Require one real chart with exact ordered data and saved backing."""
    spec = spec or {}
    expected = expected_insert_contract(spec)

    def verify_com() -> VerificationResult:
        gold = read_com(str(gold_path))
        pred = read_com(str(pred_path))
        gold_failures = com_insert_failures(gold, expected)
        if gold_failures:
            return error_result(
                "Gold chart evidence is missing or contradicts the test contract",
                {"failures": gold_failures, "gold": gold, "expected": expected},
            )
        failures = com_insert_failures(pred, expected)
        return VerificationResult(
            passed=not failures,
            message=(
                "PASS: One real chart has the requested type and ordered data"
                if not failures
                else f"FAIL: Chart insertion differs ({'; '.join(failures)})"
            ),
            details={
                "failures": failures,
                "gold": gold,
                "pred": pred,
                "expected": expected,
            },
        )

    return combine_evidence(
        verify_com,
        lambda: verify_saved(pred_path, gold_path, spec),
    )
