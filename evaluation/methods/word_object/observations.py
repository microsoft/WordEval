"""Required evidence validation without treating zero or observed absence as failure.

Existing readers return property dictionaries. ObservedProperties preserves that
interface while retaining read failures separately from successfully observed None
(for example no shading). Only a route can decide which observations are required.
"""

from math import isfinite
from typing import Any, Callable, Iterable, Mapping

from ..base import VerificationResult


class ObservedProperties(dict):
    """Property values plus failed-read diagnostics, never scoring properties."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.unavailable: dict[str, str] = {}
        self.mixed: set[str] = set()

    def record_unavailable(self, name: str, error: Exception | str):
        """Retain the cause when a parent object prevents a property read."""
        self[name] = None
        self.unavailable[name] = str(error)
        self.mixed.discard(name)

    def read(self, name: str, getter: Callable[[], Any], normalize=lambda value: value):
        try:
            raw = getter()
            if raw in (9999999, -9999999) or (normalize is bool and raw == -2):
                # Word can positively observe nonuniform formatting on a range.
                # Preserve that fact instead of bool(wdUndefined) becoming True.
                self[name] = raw
                self.mixed.add(name)
                self.unavailable.pop(name, None)
                return raw
            if isinstance(raw, (int, float)) and not isinstance(raw, bool):
                if raw in (9999999, -9999999, -2147483648) or not isfinite(raw):
                    raise ValueError(f"undefined or mixed Word property: {raw}")
            # A missing scalar must not become False/0 through normalization.
            self[name] = None if raw is None else normalize(raw)
            self.unavailable.pop(name, None)
            self.mixed.discard(name)
        except Exception as exc:
            self.record_unavailable(name, exc)
        return self[name]


def require_observations(
    states: Mapping[str, Mapping],
    required: Iterable[str],
    *,
    context: str,
    nullable: Iterable[str] = (),
) -> VerificationResult | None:
    """Reject missing/unavailable required fields before value comparison.

    Empty strings, zero and False are observations, not read failures. None is
    allowed only when the route names an absent/inactive state explicitly. A
    recorded read failure is never an observed None, including nullable fields.
    """
    nullable = set(nullable)
    required = sorted(set(required))
    for owner, state in states.items():
        if not isinstance(state, Mapping):
            return VerificationResult.error_result(
                f"Unavailable {owner} projection for {context}",
                observation_owner=owner, observation_context=context)
        failures = {}
        read_errors = getattr(state, "unavailable", {})
        mixed_fields = getattr(state, "mixed", ())
        for name in required:
            value = state.get(name)
            if name in read_errors:
                failures[name] = read_errors[name]
            elif name in mixed_fields:
                failures[name] = "required observation is mixed"
            elif name not in state or (value is None and name not in nullable):
                failures[name] = "required observation is missing"
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                if value in (9999999, -9999999, -2147483648) or not isfinite(value):
                    failures[name] = "required observation is undefined or non-finite"
        if failures:
            return VerificationResult.error_result(
                f"Unavailable {owner} observations for {context}",
                observation_owner=owner,
                observation_context=context,
                unavailable=failures,
            )
    return None
