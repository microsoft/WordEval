"""Shared score aggregation helpers for WordEval results.

The evaluator writes per-case verifier statuses. This module derives the
score views used by the website/API without changing the on-disk analysis.json
format.
"""

from __future__ import annotations

from collections import defaultdict
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


VERDICT_METHODS = ("word_object", "word_revision")
SCORE_METHODS = ("word_object", "word_revision", "run")
EVALUATED_STATUSES = {"passed", "failed"}


CaseRef = Tuple[str, str, str]  # case_id, command, variant


def split_case_id(case_id: str) -> Tuple[str, str]:
    """Split a result key of the form Command/variant."""
    if "/" not in case_id:
        return case_id, ""
    command, variant = case_id.split("/", 1)
    return command, variant


def expected_cases_from_test_cases(
    test_cases_data: Dict[str, Any],
    include_keys: Optional[Iterable[str]] = None,
) -> List[CaseRef]:
    """Return expected case refs from test_cases.json.

    If include_keys is provided, only those keys are returned. Keys not present
    in test_cases.json are parsed and included as a fallback so historical
    result files can still be scored.
    """
    include_set = set(include_keys) if include_keys is not None else None
    refs: List[CaseRef] = []
    seen = set()

    for command, variants in (test_cases_data.get("commands") or {}).items():
        for variant in (variants or {}).keys():
            case_id = f"{command}/{variant}"
            if include_set is not None and case_id not in include_set:
                continue
            refs.append((case_id, command, variant))
            seen.add(case_id)

    if include_set is not None:
        for case_id in sorted(include_set - seen):
            command, variant = split_case_id(case_id)
            refs.append((case_id, command, variant))

    return refs


def expected_cases_from_result_keys(result_keys: Iterable[str]) -> List[CaseRef]:
    """Build expected case refs directly from result keys."""
    refs: List[CaseRef] = []
    for case_id in sorted(set(result_keys)):
        command, variant = split_case_id(case_id)
        refs.append((case_id, command, variant))
    return refs


def _case_status(case_data: Dict[str, Any], method_name: str) -> str:
    if method_name == "run":
        methods = case_data.get("methods", {}) or {}
        for verifier in VERDICT_METHODS:
            if (methods.get(verifier, {}) or {}).get("status") in EVALUATED_STATUSES:
                return "passed"
        return "missing"
    return ((case_data.get("methods", {}) or {}).get(method_name, {}) or {}).get("status", "missing")


def _empty_case_stats(expected: int) -> Dict[str, Any]:
    return {
        "passed": 0,
        "failed": 0,
        "evaluated": 0,
        "expected": expected,
        "pass_rate": 0.0,          # passed / evaluated
        "strict_pass_rate": 0.0,   # passed / expected
        "coverage_rate": 0.0,      # evaluated / expected
    }


def _finalize_case_stats(stats: Dict[str, Any]) -> Dict[str, Any]:
    evaluated = stats["evaluated"]
    expected = stats["expected"]
    stats["pass_rate"] = round(stats["passed"] / evaluated * 100, 1) if evaluated else 0.0
    stats["strict_pass_rate"] = round(stats["passed"] / expected * 100, 1) if expected else 0.0
    stats["coverage_rate"] = round(evaluated / expected * 100, 1) if expected else 0.0
    return stats


def aggregate_score_bases(
    results: Dict[str, Dict[str, Any]],
    expected_cases: Sequence[CaseRef],
    include_per_command: bool = False,
) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """Aggregate case-micro and command-macro scores for each verifier.

    Case level is strict by default for leaderboard use: passed / expected.
    Command level gives every command equal weight and every variant equal
    weight only inside its parent command.
    """
    expected_cases = list(expected_cases)
    expected_count = len(expected_cases)
    by_command: Dict[str, List[str]] = defaultdict(list)
    for case_id, command, _variant in expected_cases:
        by_command[command].append(case_id)

    case_scores: Dict[str, Dict[str, Any]] = {}
    command_scores: Dict[str, Dict[str, Any]] = {}

    for method_name in SCORE_METHODS:
        case_stats = _empty_case_stats(expected_count)
        command_score_sum = 0.0
        fully_passed_commands = 0
        command_rows = []

        for command, case_ids in sorted(by_command.items()):
            command_expected = len(case_ids)
            command_passed = 0
            command_evaluated = 0
            for case_id in case_ids:
                status = _case_status(results.get(case_id, {}) or {}, method_name)
                if status in EVALUATED_STATUSES:
                    case_stats["evaluated"] += 1
                    command_evaluated += 1
                    if status == "passed":
                        case_stats["passed"] += 1
                        command_passed += 1
                    else:
                        case_stats["failed"] += 1

            command_score = command_passed / command_expected if command_expected else 0.0
            command_score_sum += command_score
            if command_expected and command_passed == command_expected:
                fully_passed_commands += 1
            command_rows.append({
                "command": command,
                "passed": command_passed,
                "evaluated": command_evaluated,
                "expected": command_expected,
                "score": round(command_score * 100, 1),
                "coverage_rate": round(command_evaluated / command_expected * 100, 1) if command_expected else 0.0,
            })

        command_count = len(by_command)
        command_pass_rate = command_score_sum / command_count * 100 if command_count else 0.0
        command_scores[method_name] = {
            "commands": command_count,
            "score_sum": round(command_score_sum, 6),
            "pass_rate": round(command_pass_rate, 1),
            "strict_pass_rate": round(command_pass_rate, 1),
            "fully_passed": fully_passed_commands,
            "expected_cases": expected_count,
            "passed_cases": case_stats["passed"],
            "evaluated_cases": case_stats["evaluated"],
            "coverage_rate": round(case_stats["evaluated"] / expected_count * 100, 1) if expected_count else 0.0,
        }
        if include_per_command:
            command_scores[method_name]["per_command"] = command_rows
        case_scores[method_name] = _finalize_case_stats(case_stats)

    return {
        "case": case_scores,
        "command": command_scores,
    }
