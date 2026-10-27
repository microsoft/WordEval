"""Visible gradient color comparison shared by shape extraction."""

from typing import Any, Dict, List


def _ole_rgb_channels(value: int) -> tuple[int, int, int]:
    color = int(value) & 0xFFFFFF
    return color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF


def _has_visible_gradient_progression(stops: List[Dict[str, Any]]) -> bool:
    """Check for a visible within-fill color progression without RGB matching."""
    usable = [
        row for row in stops if row.get("rgb") is not None and row.get("position") is not None
    ]
    if len(usable) < 2:
        return False
    if max(row["position"] for row in usable) - min(row["position"] for row in usable) < 0.5:
        return False
    for left_index, left in enumerate(usable):
        left_channels = _ole_rgb_channels(left["rgb"])
        for right in usable[left_index + 1 :]:
            right_channels = _ole_rgb_channels(right["rgb"])
            deltas = [abs(a - b) for a, b in zip(left_channels, right_channels)]
            if max(deltas) >= 12 and sum(deltas) >= 24:
                return True
    return False
