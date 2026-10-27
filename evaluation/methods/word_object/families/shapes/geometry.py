"""families / shapes / geometry with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import (
    PRESET_POSITION_THRESHOLD,
    PRESET_VALUE_MAP,
    WRAP_POLYGON_COORD_MAX,
    WRAP_POLYGON_POINT_TOLERANCE,
)
from evaluation.methods.word_object.families.shapes import (
    extract_saved as families_shapes_extract_saved,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import math


def match_shapes_by_name(
    pred_shapes: List[Dict[str, Any]], gold_shapes: List[Dict[str, Any]]
) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Match stable object identities rather than mutable XML/z-order."""
    pred_names = [shape.get("name") for shape in pred_shapes]
    gold_names = [shape.get("name") for shape in gold_shapes]
    if (
        None not in pred_names
        and None not in gold_names
        and len(set(pred_names)) == len(pred_names)
        and set(pred_names) == set(gold_names)
    ):
        pred_by_name = {shape["name"]: shape for shape in pred_shapes}
        gold_by_name = {shape["name"]: shape for shape in gold_shapes}
        names = sorted(gold_by_name)
        return (
            [pred_by_name[name] for name in names],
            [gold_by_name[name] for name in names],
        )
    return pred_shapes, gold_shapes


def closed_polygon_core(points: List[tuple[int, int]]) -> List[tuple[int, int]]:
    if len(points) > 1 and points[0] == points[-1]:
        return points[:-1]
    return points


def cyclic_polygon_distance(
    gold: List[tuple[int, int]], pred: List[tuple[int, int]]
) -> Optional[int]:
    """Return the best max-coordinate delta across rotations/directions."""
    gold_core = closed_polygon_core(gold)
    pred_core = closed_polygon_core(pred)
    if not gold_core or len(gold_core) != len(pred_core):
        return None
    best: Optional[int] = None
    for candidate in (pred_core, list(reversed(pred_core))):
        for offset in range(len(candidate)):
            rotated = candidate[offset:] + candidate[:offset]
            distance = max(
                max(abs(gx - px), abs(gy - py)) for (gx, gy), (px, py) in zip(gold_core, rotated)
            )
            best = distance if best is None else min(best, distance)
    return best


def wrap_polygon_failures(pred_path: Path, gold_path: Path) -> List[str]:
    gold = families_shapes_extract_saved.extract_wrap_polygon_signatures(gold_path)
    pred = families_shapes_extract_saved.extract_wrap_polygon_signatures(pred_path)
    if not gold:
        return []
    if len(pred) != len(gold):
        return [f"wrap polygon count: pred={len(pred)}, gold={len(gold)}"]

    def cross(a, b, c):
        return (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])

    def segments_intersect(a, b, c, d, epsilon=1e-9):
        def on_segment(p, q, r):
            return (
                min(p[0], r[0]) - epsilon <= q[0] <= max(p[0], r[0]) + epsilon
                and min(p[1], r[1]) - epsilon <= q[1] <= max(p[1], r[1]) + epsilon
            )

        values = (cross(a, b, c), cross(a, b, d), cross(c, d, a), cross(c, d, b))
        if values[0] * values[1] < -epsilon and values[2] * values[3] < -epsilon:
            return True
        return any(
            (
                abs(values[0]) <= epsilon and on_segment(a, c, b),
                abs(values[1]) <= epsilon and on_segment(a, d, b),
                abs(values[2]) <= epsilon and on_segment(c, a, d),
                abs(values[3]) <= epsilon and on_segment(c, b, d),
            )
        )

    def semantic_errors(signature: Dict[str, Any]) -> List[str]:
        if signature.get("wrap_kind") != "wrapTight":
            return [f"wrap kind is {signature.get('wrap_kind')}, expected wrapTight"]
        raw_points = signature.get("points") or []
        if any(
            x is None or y is None or not math.isfinite(float(x)) or not math.isfinite(float(y))
            for x, y in raw_points
        ):
            return ["wrap polygon has invalid coordinates"]
        # `wp:wrapPolygon` points are relative wrap coordinates bounded by
        # WRAP_POLYGON_COORD_MAX.  Word does not honour a contour written
        # outside that domain: it falls back to the bounding box, so text
        # keeps flowing exactly as it did before the tight wrap was
        # requested even though the package still carries `wrapTight` and a
        # well-formed concave point list.  Requiring the effective domain
        # keeps every legitimate Word serialization and every differently
        # proportioned in-domain contour valid.
        limit = WRAP_POLYGON_COORD_MAX + WRAP_POLYGON_POINT_TOLERANCE
        out_of_domain = [
            (x, y) for x, y in raw_points if abs(float(x)) > limit or abs(float(y)) > limit
        ]
        if out_of_domain:
            extent = max(max(abs(float(x)), abs(float(y))) for x, y in raw_points)
            return [
                "wrap polygon coordinates fall outside the effective "
                f"±{WRAP_POLYGON_COORD_MAX} relative wrap domain "
                f"(extent {extent:g}); Word renders rectangular flow"
            ]
        points = closed_polygon_core(raw_points)
        while len(points) > 1 and points[-1] == points[-2]:
            points.pop()
        if len(set(points)) < 5:
            return ["wrap polygon has fewer than five unique points"]
        min_x = min(point[0] for point in points)
        max_x = max(point[0] for point in points)
        min_y = min(point[1] for point in points)
        max_y = max(point[1] for point in points)
        span_x = max_x - min_x
        span_y = max_y - min_y
        if span_x <= 0 or span_y <= 0:
            return ["wrap polygon has a zero-area bounding box"]
        normalized = [((x - min_x) / span_x, (y - min_y) / span_y) for x, y in points]
        area = (
            abs(
                sum(
                    normalized[index][0] * normalized[(index + 1) % len(normalized)][1]
                    - normalized[(index + 1) % len(normalized)][0] * normalized[index][1]
                    for index in range(len(normalized))
                )
            )
            / 2.0
        )
        if area < 0.05:
            return ["wrap polygon enclosed area is degenerate"]
        for first in range(len(normalized)):
            a = normalized[first]
            b = normalized[(first + 1) % len(normalized)]
            for second in range(first + 1, len(normalized)):
                if (
                    second == first
                    or second == (first + 1) % len(normalized)
                    or first == (second + 1) % len(normalized)
                ):
                    continue
                c = normalized[second]
                d = normalized[(second + 1) % len(normalized)]
                if segments_intersect(a, b, c, d):
                    return ["wrap polygon self-intersects"]
        turns = [
            cross(
                normalized[index - 1],
                normalized[index],
                normalized[(index + 1) % len(normalized)],
            )
            for index in range(len(normalized))
        ]
        positive = any(value > 1e-9 for value in turns)
        negative = any(value < -1e-9 for value in turns)
        if not (positive and negative) or area > 0.9:
            return ["wrap polygon does not define a meaningful concave flow"]
        return []

    failures = []
    for index, (pred_sig, gold_sig) in enumerate(zip(pred, gold), start=1):
        if pred_sig.get("wrap_kind") != gold_sig.get("wrap_kind"):
            failures.append(
                f'shape {index} wrap kind pred={pred_sig.get("wrap_kind")}, '
                f'gold={gold_sig.get("wrap_kind")}'
            )
            continue
        failures.extend(f"shape {index} {error}" for error in semantic_errors(pred_sig))
    return failures


def bevel_render_failures(pred_path: Path, gold_path: Path) -> List[Dict]:
    gold = families_shapes_extract_saved.extract_bevel_render_signatures(gold_path)
    pred = families_shapes_extract_saved.extract_bevel_render_signatures(pred_path)
    if len(gold) != len(pred):
        return [
            {
                "property": "saved_bevel_count",
                "gold": len(gold),
                "pred": len(pred),
            }
        ]
    failures = []
    for index, (gold_sig, pred_sig) in enumerate(zip(gold, pred), start=1):
        if gold_sig.get("bevel_prst") != pred_sig.get("bevel_prst"):
            failures.append(
                {
                    "property": "bevel_prst",
                    "shape_index": index,
                    "gold": gold_sig.get("bevel_prst"),
                    "pred": pred_sig.get("bevel_prst"),
                }
            )
        for prop in (
            "scene_present",
            "camera_present",
            "light_rig_present",
            "camera_prst_present",
            "light_rig_type_present",
            "light_rig_direction_present",
        ):
            if gold_sig.get(prop) and not pred_sig.get(prop):
                failures.append(
                    {
                        "property": f"bevel_{prop}",
                        "shape_index": index,
                        "gold": True,
                        "pred": False,
                    }
                )
    return failures


def relative_layout_preservation_failures(
    pred_shapes: List[Dict[str, Any]], gold_shapes: List[Dict[str, Any]], tolerance: float
) -> List[str]:
    """Check object identity/state while deliberately ignoring position."""
    failures = []
    numeric = ("width", "height", "rotation")
    exact = (
        "name",
        "auto_shape_type",
        "is_group",
        "z_order",
        "flip_horizontal",
        "flip_vertical",
    )
    for index, (pred, gold) in enumerate(zip(pred_shapes, gold_shapes), start=1):
        for key in numeric:
            pred_value = pred.get(key)
            gold_value = gold.get(key)
            if pred_value is None and gold_value is None:
                continue
            if (
                pred_value is None
                or gold_value is None
                or abs(float(pred_value) - float(gold_value)) > tolerance
            ):
                failures.append(
                    f"shape {index} preserved {key}: " f"pred={pred_value}, gold={gold_value}"
                )
        for key in exact:
            pred_value = pred.get(key)
            gold_value = gold.get(key)
            if (pred_value is not None or gold_value is not None) and pred_value != gold_value:
                failures.append(
                    f"shape {index} preserved {key}: " f"pred={pred_value}, gold={gold_value}"
                )
    return failures


def normalized_shape_start(shape: Dict[str, Any], axis: str) -> float:
    """Return an unrotated shape-frame start in page coordinates."""
    if axis == "horizontal":
        return calculate_actual_position(
            float(shape["left"]),
            float(shape["width"]),
            int(shape["relative_h_position"]),
            float(shape.get("page_width", 612)),
            float(shape.get("left_margin", 72)),
            float(shape.get("right_margin", 72)),
        )
    return calculate_actual_position(
        float(shape["top"]),
        float(shape["height"]),
        int(shape["relative_v_position"]),
        float(shape.get("page_height", 792)),
        float(shape.get("top_margin", 72)),
        float(shape.get("bottom_margin", 72)),
    )


def normalized_alignment_value(shape: Dict[str, Any], alignment_type: str) -> float:
    horizontal = alignment_type in {"left", "right", "center_h"}
    axis = "horizontal" if horizontal else "vertical"
    start = normalized_shape_start(shape, axis)
    size = float(shape["width" if horizontal else "height"])
    if alignment_type in {"right", "bottom"}:
        return start + size
    if alignment_type in {"center_h", "center_v"}:
        return start + size / 2
    return start


def orthogonal_alignment_preservation_failures(
    pred_shapes: List[Dict[str, Any]],
    gold_shapes: List[Dict[str, Any]],
    alignment_type: str,
    tolerance: float,
) -> List[str]:
    """Reject drift on the axis an alignment command does not target."""
    target_is_horizontal = alignment_type in {"left", "right", "center_h"}
    orthogonal_axis = "vertical" if target_is_horizontal else "horizontal"
    failures = []
    for index, (pred, gold) in enumerate(zip(pred_shapes, gold_shapes), start=1):
        pred_value = normalized_shape_start(pred, orthogonal_axis)
        gold_value = normalized_shape_start(gold, orthogonal_axis)
        if abs(pred_value - gold_value) > tolerance:
            failures.append(
                f"shape {index} orthogonal {orthogonal_axis} position: "
                f"pred={pred_value:.2f}, gold={gold_value:.2f}"
            )
    return failures


def is_preset_position(value: float) -> bool:
    """Check if a position value is a Word preset constant."""
    return value < PRESET_POSITION_THRESHOLD


def calculate_actual_position(
    raw_pos: float,
    shape_size: float,
    rel_position: int,
    page_size: float,
    margin_start: float,
    margin_end: float,
) -> float:
    """
    Calculate actual rendered position from Word's position values.

    Args:
        raw_pos: Raw position value from shape.Left or shape.Top
        shape_size: Shape width or height
        rel_position: 0=margin-relative, 1=page-relative
        page_size: Page width or height
        margin_start: Left or top margin
        margin_end: Right or bottom margin

    Returns:
        Actual position in points from page edge
    """
    # Check if this is a preset constant (large negative value)
    if is_preset_position(raw_pos):
        # Determine the reference area based on relative position
        if rel_position == 0:  # margin-relative
            base = margin_start
            area_size = page_size - margin_start - margin_end
        else:  # page-relative
            base = 0
            area_size = page_size

        # Map preset constant to position type
        pos_type = PRESET_VALUE_MAP.get(raw_pos, "unknown")

        if pos_type == "start":
            return base
        elif pos_type == "center":
            return base + (area_size - shape_size) / 2
        elif pos_type == "end":
            return base + area_size - shape_size
        else:
            # Unknown preset, return raw value
            return raw_pos
    else:
        # Not a preset - raw_pos is an offset from the reference point.
        # Word may report AddShape output relative to the text column or
        # anchor paragraph (2). In our single-column benchmark documents,
        # that origin is the page text area, so add the page margin to
        # normalize to page coordinates. This makes column-relative 28pt
        # comparable to page-relative 100pt when the margin is 72pt.
        if rel_position in (0, 2):  # margin/column/paragraph-relative
            return margin_start + raw_pos
        else:  # page-relative
            return raw_pos
