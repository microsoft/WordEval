"""Check or refresh composition intent targets without regenerating DOCX files."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from generator.compositions.atomic_intents import (
    CONTENT_FORMATTING_ATOMIC_CASES,
    CONTENT_OPERATIONS_ATOMIC_CASES,
    OBJECT_EFFECT_ATOMIC_CASES,
    _targets_for_content_operation,
    _targets_for_object_effect_slot,
    _targets_for_slot,
    build_layout_coupled_atomic_intents,
    build_prerequisite_chain_atomic_intents,
    build_scope_changing_atomic_intents,
)
from generator.compositions.additive.content_formatting import PARAGRAPH_VARIANTS
from generator.test_sets.spec_utils import write_spec_if_changed


def _testset_from_path(path: Path) -> str:
    parts = path.resolve().parts
    try:
        benchmark_index = parts.index("benchmark")
    except ValueError as exc:
        raise ValueError(f"Spec is not under benchmark/: {path}") from exc
    return parts[benchmark_index + 1]


def _refresh_content_formatting_intents(
    intents: list[dict[str, Any]],
    target_comparisons: dict[str, dict[str, str]] | None = None,
) -> list[dict[str, Any]]:
    target_comparisons = target_comparisons or {}
    refreshed = []
    for intent in intents:
        atomic_case = intent.get("atomic_case")
        operation = intent.get("operation")
        value = next((
            value
            for (slot, value), candidate in CONTENT_FORMATTING_ATOMIC_CASES.items()
            if slot == operation and candidate == atomic_case
        ), None)
        if value is not None:
            targets = _targets_for_slot(operation, value)
        elif operation in {"indent", "list_indent"}:
            if str(atomic_case).startswith("IndentIncreaseWord/"):
                targets = _targets_for_slot(operation, "increase")
            elif str(atomic_case).startswith("IndentDecreaseWord/"):
                targets = _targets_for_slot(operation, "decrease")
            else:
                targets = intent.get("verification_targets", [])
        else:
            targets = intent.get("verification_targets", [])
        slot_comparisons = target_comparisons.get(operation, {})
        for target in targets:
            comparison = slot_comparisons.get(target.get("property"))
            if comparison is not None:
                target["comparison"] = comparison
        refreshed.append({**intent, "verification_targets": targets})
    return refreshed


def _refresh_mapped_intents(
    intents: list[dict[str, Any]],
    mapping: dict[tuple[str, Any], str],
    target_builder,
) -> list[dict[str, Any]]:
    refreshed = []
    for intent in intents:
        atomic_case = intent.get("atomic_case")
        operation = intent.get("operation")
        value = next((
            value
            for (slot, value), candidate in mapping.items()
            if slot == operation and candidate == atomic_case
        ), None)
        targets = (
            target_builder(operation, value)
            if value is not None
            else intent.get("verification_targets", [])
        )
        refreshed.append({**intent, "verification_targets": targets})
    return refreshed


def refreshed_atomic_intents(path: Path, spec: dict[str, Any]) -> list[dict]:
    composition = spec.get("composition") or {}
    current = composition.get("atomic_intents") or []
    testset = _testset_from_path(path)

    if testset == "additive_composition_x_content_formatting":
        target_comparisons = {}
        if path.parent.parent.name == "AdditiveParagraphComposition":
            variant = PARAGRAPH_VARIANTS.get(path.parent.name, {})
            target_comparisons = variant.get("verification_comparisons", {})
        return _refresh_content_formatting_intents(
            current, target_comparisons=target_comparisons
        )
    if testset == "additive_composition_x_content_operations":
        return _refresh_mapped_intents(
            current,
            CONTENT_OPERATIONS_ATOMIC_CASES,
            _targets_for_content_operation,
        )
    if testset == "additive_composition_x_object_effect":
        return _refresh_mapped_intents(
            current,
            OBJECT_EFFECT_ATOMIC_CASES,
            _targets_for_object_effect_slot,
        )

    steps = composition.get("steps") or []
    if testset == "interaction_composition_x_scope_changing":
        return build_scope_changing_atomic_intents(steps)
    if testset == "interaction_composition_x_prerequisite_chain":
        return build_prerequisite_chain_atomic_intents(steps)
    if testset == "interaction_composition_x_layout_coupled":
        return build_layout_coupled_atomic_intents(steps)
    raise ValueError(f"Unsupported composition test set: {testset}")


def main() -> None:
    parser = argparse.ArgumentParser()
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--spec", action="append", type=Path)
    source.add_argument(
        "--all-benchmark",
        action="store_true",
        help="Check every supported composition spec under benchmark/",
    )
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()

    paths = args.spec or []
    if args.all_benchmark:
        repo_root = Path(__file__).resolve().parents[2]
        benchmark = repo_root / "benchmark"
        supported_sets = (
            "additive_composition_x_content_formatting",
            "additive_composition_x_content_operations",
            "additive_composition_x_object_effect",
            "interaction_composition_x_scope_changing",
            "interaction_composition_x_prerequisite_chain",
            "interaction_composition_x_layout_coupled",
        )
        paths = [
            path
            for testset in supported_sets
            for path in sorted((benchmark / testset).glob("*/*/spec.json"))
        ]

    changed = 0
    for path in paths:
        spec = json.loads(path.read_text(encoding="utf-8"))
        expected = refreshed_atomic_intents(path, spec)
        current = spec.get("composition", {}).get("atomic_intents") or []
        if current == expected:
            print(f"OK {path}")
            continue
        changed += 1
        print(f"STALE {path}")
        if args.write:
            spec["composition"]["atomic_intents"] = expected
            write_spec_if_changed(path, spec)
    if changed and not args.write:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
