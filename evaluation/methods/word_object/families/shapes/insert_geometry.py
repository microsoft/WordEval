"""families / shapes / insert_geometry with explicit evidence dependencies."""

from typing import Any, Dict, Optional
import math


def is_preset_shape_position(value: float) -> bool:
    return value < -900000.0


def effective_shape_coordinate(shape: Dict[str, Any], axis: str) -> float:
    if axis == "left":
        raw_key, size_key, rel_key = "left", "width", "relative_h_position"
        page_key, start_key, end_key = ("page_width", "left_margin", "right_margin")
    else:
        raw_key, size_key, rel_key = "top", "height", "relative_v_position"
        page_key, start_key, end_key = ("page_height", "top_margin", "bottom_margin")
    raw_pos = shape.get(raw_key, 0) or 0
    size = shape.get(size_key, 0) or 0
    relative = shape.get(rel_key, 1)
    page_size = shape.get(page_key, 0) or 0
    margin_start = shape.get(start_key, 0) or 0
    margin_end = shape.get(end_key, 0) or 0
    if is_preset_shape_position(raw_pos):
        if relative in (0, 2):
            base = margin_start
            area_size = page_size - margin_start - margin_end
        else:
            base = 0
            area_size = page_size
        preset = {
            -999998.0: "start",
            -999999.0: "start",
            -999995.0: "center",
            -999996.0: "end",
            -999997.0: "end",
        }.get(raw_pos)
        if preset == "start":
            return base
        if preset == "center":
            return base + (area_size - size) / 2
        if preset == "end":
            return base + area_size - size
        return raw_pos
    if relative in (0, 2):
        return margin_start + raw_pos
    return raw_pos


def shape_usable_page_failure(shape: Dict[str, Any]) -> Optional[str]:
    width = float(shape.get("width") or 0)
    height = float(shape.get("height") or 0)
    if width < 2.0 or height < 2.0:
        return f"degenerate target (width={width}, height={height})"
    if shape.get("visible") is False:
        return "target shape is hidden"
    fill_visible = shape.get("fill_visible")
    fill_type = shape.get("fill_type")
    fill_transparency = shape.get("fill_transparency")
    has_fill = (
        fill_visible is not False
        and fill_type not in (None, 0)
        and (fill_transparency is None or float(fill_transparency) < 0.999)
    )
    has_line = shape.get("line_visible") is True
    if not has_fill and not has_line:
        return "target shape has no visible fill or line"
    left = effective_shape_coordinate(shape, "left")
    top = effective_shape_coordinate(shape, "top")
    page_width = float(shape.get("page_width") or 0)
    page_height = float(shape.get("page_height") or 0)
    if page_width and page_height:
        center_x = left + width / 2
        center_y = top + height / 2
        angle = math.radians(float(shape.get('rotation') or 0) % 360)
        half_width = (abs(width * math.cos(angle)) + abs(height * math.sin(angle))) / 2
        half_height = (abs(width * math.sin(angle)) + abs(height * math.cos(angle))) / 2
        if not (-0.5 <= center_x-half_width and center_x+half_width <= page_width+0.5
                and -0.5 <= center_y-half_height and center_y+half_height <= page_height+0.5):
            return (
                "target shape is clipped by the usable page boundary "
                f"(left={left}, top={top}, width={width}, height={height})"
            )
    return None
