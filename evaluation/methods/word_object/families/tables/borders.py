"""Complete physical-cell borders when a table aggregate is positively mixed."""
from ....base import VerificationResult
from ...observations import ObservedProperties, require_observations
from .border_xml import border_style

CELL_EDGES = ("top", "left", "bottom", "right")
TABLE_EDGES = CELL_EDGES + ("horizontal", "vertical")


def read_borders(obj, edges):
    result = {}
    for kind, name in enumerate(edges, 1):
        edge = ObservedProperties()
        edge.read("line_style", lambda: border_style(obj, kind, name,
                  scope="table" if edges == TABLE_EDGES else "cell"))
        edge.read("line_width", lambda: obj.Borders(-kind).LineWidth if edge["line_style"] != 0 else None)
        edge.read("color", lambda: obj.Borders(-kind).Color if edge["line_style"] != 0 else None)
        result[name] = edge
    return result


def required_border_fields(edge):
    return {"line_style"} if edge.get("line_style") == 0 else {"line_style", "line_width", "color"}


def has_mixed_borders(table):
    return any(required_border_fields(edge) & set(getattr(edge, "mixed", ()))
               for edge in (table.get("borders") or {}).values())


def validate_cell_borders(table, owner):
    error = require_observations({owner: table}, {"cell_borders", "cell_contents"}, context="physical table borders")
    if error is not None:
        return error
    expected = {(row, column) for row, values in enumerate(table["cell_contents"], 1)
                for column, value in enumerate(values, 1) if value is not None}
    observed = set()
    for cell in table["cell_borders"]:
        error = require_observations({owner: cell}, {"row", "column", "borders"}, context="physical table cell borders")
        if error is not None:
            return error
        coordinate = (cell["row"], cell["column"])
        if coordinate in observed:
            return VerificationResult.error_result("Duplicate physical-cell borders", observation_owner=owner)
        observed.add(coordinate)
        for name in CELL_EDGES:
            edge = cell["borders"].get(name, {})
            error = require_observations({owner: edge}, required_border_fields(edge),
                                         context=f"physical cell {coordinate} border {name}")
            if error is not None:
                return error
    if observed != expected:
        return VerificationResult.error_result("Incomplete physical-cell borders", observation_owner=owner)
    return None


def cell_border_signature(table):
    """Use only after validation; inactive edges have no width or color contract."""
    return sorted((cell["row"], cell["column"], tuple(
        (name, tuple((field, cell["borders"][name][field])
                     for field in sorted(required_border_fields(cell["borders"][name]))))
        for name in CELL_EDGES)) for cell in table["cell_borders"])
