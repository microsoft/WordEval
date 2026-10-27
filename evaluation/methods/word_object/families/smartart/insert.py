"""families / smartart / insert with explicit evidence dependencies."""

from collections import defaultdict
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.smartart import (
    insert_saved as families_smartart_insert_saved,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re
from .placement import page_bounds_failure


def expected_smartart_insert_contract(spec: Dict[str, Any]) -> Dict[str, Any]:
    labels = list(
        spec.get("verification", {}).get("primary", {}).get("params", {}).get("labels", [])
    )
    variation = spec.get("variation", "")
    if variation == "insert_hierarchy":
        edge_map = {
            "__root__": ["Boss"],
            "Boss": ["Team1", "Team2"],
            "Team1": ["Member1", "Member2"],
            "Team2": ["Member3"],
        }
        levels = [1, 2, 3, 3, 2, 3]
        layout_token = "hierarchy"
        saved_layout_suffix = "/hierarchy1"
    else:
        edge_map = {"__root__": labels}
        levels = [1] * len(labels)
        layout_token = "basicblocklist"
        saved_layout_suffix = "/default"
    return {
        "labels": labels,
        "levels": levels,
        "edge_map": edge_map,
        "layout_token": layout_token,
        "saved_layout_suffix": saved_layout_suffix,
    }


def smartart_edge_map(edges: List[Dict[str, Any]]) -> Dict[str, List[str]]:
    grouped: Dict[str, List[tuple[int, str]]] = defaultdict(list)
    for edge in edges:
        grouped[edge.get("parent")].append((int(edge.get("order", 0)), edge.get("child")))
    return {
        parent: [child for _, child in sorted(children)] for parent, children in grouped.items()
    }


def smartart_com_contract_failures(state: Dict[str, Any], expected: Dict[str, Any]) -> List[str]:
    failures = []
    if not expected.get("labels"):
        failures.append("test contract lacks SmartArt node labels")
        return failures
    if state.get("smartart_count") != 1:
        failures.append(f"expected one SmartArt, found {state.get('smartart_count')}")
    if state.get("visual_object_count") != 1:
        failures.append(
            f"expected one body visual object, found {state.get('visual_object_count')}"
        )
    layout = re.sub(r"[^a-z0-9]", "", state.get("layout_name", "").casefold())
    if expected.get("layout_token") not in layout:
        failures.append(f"layout {state.get('layout_name')!r} is not the requested layout")
    if state.get("all_node_texts") != expected.get("labels"):
        failures.append("SmartArt node text/order differs")
    if state.get("all_node_levels") != expected.get("levels"):
        failures.append("SmartArt hierarchy levels differ")
    bounds_failure = page_bounds_failure(state.get('page_bounds'))
    if bounds_failure:
        failures.append(bounds_failure)
    return failures


def smartart_saved_contract_failures(state: Dict[str, Any], expected: Dict[str, Any]) -> List[str]:
    failures = list(state.get("errors") or [])
    if state.get("object_count") != 1:
        failures.append(f"expected one referenced SmartArt, found {state.get('object_count')}")
    if state.get("body_object_count") != 1:
        failures.append(f"expected one saved body object, found {state.get('body_object_count')}")
    diagrams = state.get("diagrams") or []
    if len(diagrams) != 1:
        return failures
    diagram = diagrams[0]
    if not (diagram.get("layout_id") or "").endswith(expected.get("saved_layout_suffix", "")):
        failures.append("saved SmartArt layout identity differs")
    if diagram.get("node_order") != expected.get("labels"):
        failures.append("saved SmartArt node text/order differs")
    if diagram.get("node_levels") != expected.get("levels"):
        failures.append("saved SmartArt hierarchy levels differ")
    if smartart_edge_map(diagram.get("edges") or []) != expected.get("edge_map"):
        failures.append("saved SmartArt parent/child graph differs")
    if diagram.get("drawing_texts") != expected.get("labels"):
        failures.append("saved SmartArt drawing cache differs")
    return failures


def verify_smartart_insert_saved(
    pred_path: Path,
    gold_path: Path,
    spec: Dict[str, Any],
    *,
    error_with_details,
    saved_part_relationships,
    verify_saved_insertion_context,
) -> VerificationResult:
    expected = expected_smartart_insert_contract(spec)
    try:
        gold = families_smartart_insert_saved.extract_saved_smartart_signature(
            gold_path, saved_part_relationships=saved_part_relationships
        )
    except Exception as exc:
        return error_with_details(f"Cannot extract gold saved SmartArt evidence: {exc}")
    gold_failures = smartart_saved_contract_failures(gold, expected)
    if gold_failures:
        return error_with_details(
            "Gold saved SmartArt evidence is incomplete or invalid",
            {"failures": gold_failures, "gold": gold, "expected": expected},
        )
    try:
        pred = families_smartart_insert_saved.extract_saved_smartart_signature(
            pred_path, saved_part_relationships=saved_part_relationships
        )
    except Exception as exc:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Predicted saved SmartArt package is unreadable ({exc})",
            details={"gold": gold, "expected": expected},
        )
    failures = smartart_saved_contract_failures(pred, expected)
    context_failure = verify_saved_insertion_context(pred_path, gold_path)
    if context_failure is not None:
        failures.append("surrounding document content/layout changed")
    return VerificationResult(
        passed=not failures,
        message=(
            "PASS: Saved SmartArt layout, graph, cache, and context are valid"
            if not failures
            else f"FAIL: Saved SmartArt insertion differs ({'; '.join(failures)})"
        ),
        details={
            "match_mode": "saved_ooxml_smartart_layout_graph_cache",
            "failures": failures,
            "gold": gold,
            "pred": pred,
            "expected": expected,
            "context_failure": (context_failure.to_dict() if context_failure else None),
        },
    )


def verify_smartart_insert(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    com_primary_with_saved_guardrail,
    error_with_details,
    read_inserted_smartart_state,
    saved_part_relationships,
    verify_saved_insertion_context,
) -> VerificationResult:
    """Require one real SmartArt with the requested layout and hierarchy."""
    spec = spec or {}
    expected = expected_smartart_insert_contract(spec)

    def verify_com() -> VerificationResult:
        gold = read_inserted_smartart_state(str(gold_path))
        pred = read_inserted_smartart_state(str(pred_path))
        gold_failures = smartart_com_contract_failures(gold, expected)
        if gold_failures:
            return error_with_details(
                "Gold SmartArt evidence is missing or contradicts the test contract",
                {"failures": gold_failures, "gold": gold, "expected": expected},
            )
        failures = smartart_com_contract_failures(pred, expected)
        return VerificationResult(
            passed=not failures,
            message=(
                "PASS: One real SmartArt has the requested layout and hierarchy"
                if not failures
                else f"FAIL: SmartArt insertion differs ({'; '.join(failures)})"
            ),
            details={
                "failures": failures,
                "gold": gold,
                "pred": pred,
                "expected": expected,
            },
        )

    return com_primary_with_saved_guardrail(
        verify_com,
        lambda: verify_smartart_insert_saved(
            pred_path,
            gold_path,
            spec,
            error_with_details=error_with_details,
            saved_part_relationships=saved_part_relationships,
            verify_saved_insertion_context=verify_saved_insertion_context,
        ),
    )
