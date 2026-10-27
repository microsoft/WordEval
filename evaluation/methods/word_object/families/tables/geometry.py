"""Physical-cell geometry when Word cannot enumerate merged rows/columns."""
from ....base import VerificationResult
from ...observations import require_observations


def validate_cell_geometry(table, owner):
    cells = table["cell_geometry"]
    expected = {(row, column) for row, values in enumerate(table["cell_contents"], 1)
                for column, value in enumerate(values, 1) if value is not None}
    observed = set()
    for cell in cells:
        error = require_observations({owner: cell}, {"row", "column", "width", "height_rule", "height"},
            context="physical table cell geometry", nullable={"height"} if cell.get("height_rule") == 0 else ())
        if error is not None:
            return error
        coordinate = (cell["row"], cell["column"])
        if coordinate in observed or cell["width"] <= 0 or cell["height_rule"] not in {0, 1, 2}:
            return VerificationResult.error_result("Invalid physical-cell geometry", observation_owner=owner)
        observed.add(coordinate)
    if observed != expected:
        return VerificationResult.error_result("Incomplete physical-cell geometry", observation_owner=owner)
    return None


def physical_geometry_differences(gold, pred):
    if len(gold) != len(pred):
        return [{"property": "physical_cell_count", "gold": len(gold), "pred": len(pred)}]
    differences = []
    for index, (g, p) in enumerate(zip(gold, pred)):
        for field in ("row", "column", "width", "height_rule", "height"):
            gv, pv = g[field], p[field]
            if field in {"width", "height"} and gv is not None and pv is not None:
                equal = abs(gv - pv) <= 2.0
            else:
                equal = gv == pv
            if not equal:
                differences.append({"property": f"physical_cell_{index}_{field}", "gold": gv, "pred": pv})
    return differences
