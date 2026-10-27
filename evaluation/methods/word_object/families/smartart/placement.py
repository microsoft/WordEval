"""Resolve the actual SmartArt frame against its containing Word page."""
import math


def read_page_bounds(doc, shape, source):
    doc.Repaginate()
    anchor = shape.Range if source == 'inline' else shape.Anchor
    setup = anchor.Sections(1).PageSetup
    width, height = float(shape.Width), float(shape.Height)
    page_width, page_height = float(setup.PageWidth), float(setup.PageHeight)

    def information(code):
        value = float(anchor.Information(code))
        if value == -1:
            raise RuntimeError(f'SmartArt page coordinate {code} is unavailable')
        return value

    if source == 'inline':
        left, top = information(5), information(6)
        angle = 0.0
        visible = True
    else:
        def coordinate(axis):
            horizontal = axis == 'horizontal'
            relative = int(shape.RelativeHorizontalPosition if horizontal else shape.RelativeVerticalPosition)
            raw = float(shape.Left if horizontal else shape.Top)
            size, page = (width, page_width) if horizontal else (height, page_height)
            start = float(setup.LeftMargin if horizontal else setup.TopMargin)
            end = float(setup.RightMargin if horizontal else setup.BottomMargin)
            if relative == 0:
                origin, span = start, page - start - end
            elif relative == 1:
                origin, span = 0.0, page
            elif relative == 2 and horizontal:
                if int(setup.TextColumns.Count) == 1:
                    origin, span = start, page - start - end
                else:
                    # Word exposes the current column boundary independently
                    # from the character's indentation within that column.
                    origin = information(5) - information(7)
                    if raw < -900000:
                        raise RuntimeError('Aligned SmartArt in a multi-column section is unresolved')
                    span = 0.0
            elif relative in (2, 3):
                origin = information(5 if horizontal else 6)
                if raw < -900000:
                    raise RuntimeError('Aligned character/paragraph SmartArt position is unresolved')
                span = 0.0
            elif relative == 4:
                origin, span = 0.0, start
            elif relative == 5:
                origin, span = page - end, end
            else:
                raise RuntimeError(f'Unresolved SmartArt {axis} reference {relative}')
            if raw >= -900000:
                return origin + raw
            if raw in (-999998.0, -999999.0):return origin
            if raw == -999995.0:return origin + (span - size) / 2
            if raw in (-999996.0, -999997.0):return origin + span - size
            raise RuntimeError(f'Unresolved SmartArt alignment sentinel {raw}')
        left, top = coordinate('horizontal'), coordinate('vertical')
        angle = math.radians(float(shape.Rotation) % 360)
        visible = bool(shape.Visible)
    rotated_width = abs(width * math.cos(angle)) + abs(height * math.sin(angle))
    rotated_height = abs(width * math.sin(angle)) + abs(height * math.cos(angle))
    left -= (rotated_width - width) / 2
    top -= (rotated_height - height) / 2
    return {'left':left,'top':top,'right':left+rotated_width,'bottom':top+rotated_height,
            'page_width':page_width,'page_height':page_height,'visible':visible}


def page_bounds_failure(bounds):
    required = ('left','top','right','bottom','page_width','page_height','visible')
    if not isinstance(bounds, dict) or any(key not in bounds for key in required):
        return 'SmartArt usable-page bounds were not observed'
    if not bounds['visible']:
        return 'SmartArt is hidden'
    if not all(math.isfinite(float(bounds[key])) for key in required[:-1]):
        return 'SmartArt geometry is non-finite'
    if bounds['right'] <= bounds['left'] or bounds['bottom'] <= bounds['top']:
        return 'SmartArt has no usable frame'
    if (bounds['left'] < -0.5 or bounds['top'] < -0.5
            or bounds['right'] > bounds['page_width'] + 0.5
            or bounds['bottom'] > bounds['page_height'] + 0.5):
        return 'SmartArt frame is clipped by the page boundary'
    return None
