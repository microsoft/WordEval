"""Helpers for composition specs shared by Word COM verifiers."""

from typing import Any, Dict, Iterable, List, Optional

from ..base import VerificationResult


def _dedupe(values: Iterable[Optional[str]]) -> List[str]:
    return list(dict.fromkeys(value for value in values if value))


def atomic_case_command_id(atomic_case: Optional[str]) -> Optional[str]:
    """Return the command/idMso prefix from an atomic case key."""
    if not atomic_case:
        return None
    return str(atomic_case).split("/", 1)[0]


def composition_command_ids(spec: Optional[Dict[str, Any]]) -> List[str]:
    """Return composition command ids from current or legacy composition specs."""
    composition = (spec or {}).get("composition", {}) or {}
    intents = composition.get("atomic_intents") or []
    command_ids = _dedupe(
        [
            atomic_case_command_id(intent.get("atomic_case"))
            for intent in intents
            if isinstance(intent, dict)
        ]
    )
    if command_ids:
        return command_ids

    legacy_commands = composition.get("commands") or []
    return _dedupe(
        [
            (
                item.get("idMso") or item.get("command") or item.get("id")
                if isinstance(item, dict)
                else str(item)
            )
            for item in legacy_commands
        ]
    )


def composition_operation_ids(spec: Optional[Dict[str, Any]]) -> List[str]:
    """Return operation names declared by atomic intents and step specs."""
    composition = (spec or {}).get("composition", {}) or {}
    operations: List[str] = []
    for intent in composition.get("atomic_intents") or []:
        if isinstance(intent, dict):
            operations.append(intent.get("operation"))
    for step in composition.get("steps") or []:
        if isinstance(step, dict):
            operations.append(step.get("operation"))
    return _dedupe(operations)


def composition_steps(spec: Optional[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Return normalized composition steps from explicit steps or atomic_intents."""
    composition = (spec or {}).get("composition", {}) or {}
    steps = composition.get("steps") or []
    if steps:
        return [step for step in steps if isinstance(step, dict)]

    intent_steps = []
    for intent in composition.get("atomic_intents") or []:
        if not isinstance(intent, dict):
            continue
        step: Dict[str, Any] = {}
        command_id = atomic_case_command_id(intent.get("atomic_case"))
        if command_id:
            step["idMso"] = command_id
        operation = intent.get("operation")
        if operation:
            step["operation"] = operation
        if step:
            intent_steps.append(step)
    if intent_steps:
        return intent_steps

    legacy_steps = []
    for command in composition.get("commands") or []:
        if isinstance(command, dict):
            step = dict(command)
            command_id = step.get("idMso") or step.get("command") or step.get("id")
            if command_id:
                step["idMso"] = command_id
            if step:
                legacy_steps.append(step)
        elif command:
            legacy_steps.append({"idMso": str(command)})
    if legacy_steps:
        return legacy_steps

    return []


def composition_step_key(
    route: str,
    step_index: int,
    step: Dict[str, Any],
    seen: set[str],
) -> tuple[Optional[str], Optional[VerificationResult]]:
    """Validate one declared composition step before dispatching it."""
    operation = step.get("operation")
    command = step.get("idMso")
    if not operation and not command:
        result = VerificationResult.error_result(
            f"{route} composition step {step_index} has neither operation " "nor idMso"
        )
        result.details.update(
            {
                "route": route,
                "step_index": step_index,
                "step": step,
                "reason": "composition_step_missing_dispatch_key",
            }
        )
        return None, result

    key = str(operation or command)
    if key in seen:
        result = VerificationResult.error_result(
            f"{route} composition step {step_index} duplicates {key!r}"
        )
        result.details.update(
            {
                "route": route,
                "step_index": step_index,
                "step": step,
                "duplicate_key": key,
                "reason": "duplicate_composition_step",
            }
        )
        return None, result
    seen.add(key)
    return key, None


def composition_step_declaration_error(
    route: str,
    spec: Optional[Dict[str, Any]],
) -> Optional[VerificationResult]:
    """Reject a declared step list that normalization would drop."""
    composition = (spec or {}).get("composition") or {}
    operation_count_present = "operation_count" in composition
    operation_count = composition.get("operation_count")
    if operation_count_present and (type(operation_count) is not int or operation_count < 0):
        result = VerificationResult.error_result(f"{route} composition has invalid operation_count")
        result.details.update(
            {
                "route": route,
                "operation_count": operation_count,
                "reason": "invalid_composition_operation_count",
            }
        )
        return result

    step_fields = ("steps", "atomic_intents", "commands")
    present_fields = [field for field in step_fields if field in composition]
    malformed_fields = [
        field for field in present_fields if not isinstance(composition.get(field), list)
    ]
    if malformed_fields:
        result = VerificationResult.error_result(
            f"{route} composition has malformed step declaration"
        )
        result.details.update(
            {
                "route": route,
                "declared_fields": present_fields,
                "malformed_fields": malformed_fields,
                "reason": "malformed_composition_step_declaration",
            }
        )
        return result

    declared_field = next(
        (field for field in step_fields if composition.get(field)),
        None,
    )
    if declared_field is None:
        if not present_fields and (not operation_count_present or operation_count == 0):
            return None
        result = VerificationResult.error_result(
            f"{route} composition declares no usable verification steps"
        )
        result.details.update(
            {
                "route": route,
                "declared_fields": present_fields,
                "operation_count": composition.get("operation_count"),
                "reason": "empty_composition_step_declaration",
            }
        )
        return result

    declared_steps: Any = composition[declared_field]

    invalid_indexes = []
    if not isinstance(declared_steps, list):
        invalid_indexes = ["not_a_list"]
    elif declared_field == "commands":
        invalid_indexes = [
            index
            for index, item in enumerate(declared_steps, start=1)
            if not item
            or (
                isinstance(item, dict)
                and not any(
                    item.get(key)
                    for key in (
                        "operation",
                        "idMso",
                        "command",
                        "id",
                    )
                )
            )
        ]
    elif declared_field == "atomic_intents":
        invalid_indexes = [
            index
            for index, item in enumerate(declared_steps, start=1)
            if not isinstance(item, dict)
            or not item
            or (not item.get("operation") and not atomic_case_command_id(item.get("atomic_case")))
            or (
                item.get("atomic_case") is not None
                and not atomic_case_command_id(item.get("atomic_case"))
            )
        ]
    else:
        invalid_indexes = [
            index
            for index, item in enumerate(declared_steps, start=1)
            if not isinstance(item, dict)
        ]

    if invalid_indexes:
        result = VerificationResult.error_result(
            f"{route} composition has malformed {declared_field}"
        )
        result.details.update(
            {
                "route": route,
                "declared_field": declared_field,
                "invalid_step_indexes": invalid_indexes,
                "reason": "malformed_composition_step_declaration",
            }
        )
        return result

    if operation_count_present and operation_count != len(declared_steps):
        result = VerificationResult.error_result(
            f"{route} composition operation_count does not match " f"{declared_field}"
        )
        result.details.update(
            {
                "route": route,
                "declared_field": declared_field,
                "operation_count": operation_count,
                "declared_step_count": len(declared_steps),
                "reason": "composition_operation_count_mismatch",
            }
        )
        return result
    return None


def unhandled_composition_step(
    route: str,
    step_index: int,
    step: Dict[str, Any],
) -> VerificationResult:
    result = VerificationResult.error_result(
        f"Unsupported {route} composition step {step_index}: "
        f'operation={step.get("operation")!r}, idMso={step.get("idMso")!r}'
    )
    result.details.update(
        {
            "route": route,
            "step_index": step_index,
            "step": step,
            "reason": "unsupported_composition_step",
        }
    )
    return result
