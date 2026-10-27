"""Validate route-owned table evidence before semantic comparisons."""
from ...observations import require_observations
from ....base import VerificationResult
from .geometry import validate_cell_geometry
from .borders import validate_cell_borders


TABLE_FIELDS = {
    "row_count", "column_count", "cell_contents", "cell_alignments",
    "cell_text_orientations", "cell_paragraph_alignments", "style",
    "borders", "column_widths", "row_heights",
}


def validate_tables(gold, pred, required, *, allow_empty=False, outer_border_style_only=False, physical_borders=False):
    if gold is None or pred is None or (not gold and not allow_empty):
        return VerificationResult.error_result("Required table observations are unavailable")
    for owner, tables in (("gold", gold), ("pred", pred)):
        for index, table in enumerate(tables, 1):
            error = require_observations({owner: table}, required, context=f"table {index}")
            if error is not None:
                return error
            if "cell_geometry" in required:
                error = validate_cell_geometry(table, owner)
                if error is not None:
                    return error
            if "borders" in required:
                edges = ("top", "left", "bottom", "right") if outer_border_style_only else ("top", "left", "bottom", "right", "horizontal", "vertical")
                if outer_border_style_only and owner == "pred" and index <= len(gold):
                    # Gold has already been fully validated. A positively
                    # observed wrong outer style disproves the uniform-style
                    # requirement even if another edge is mixed/unavailable.
                    # Never use the undefined sentinel itself as a mismatch.
                    differences = []
                    for edge in edges:
                        observed = table["borders"].get(edge, {})
                        if require_observations({owner: observed}, {"line_style"}, context="outer border") is not None:
                            continue
                        expected = gold[index - 1]["borders"][edge]["line_style"]
                        if observed["line_style"] != expected:
                            differences.append(dict(table=index, edge=edge, gold=expected, pred=observed["line_style"]))
                    if differences:
                        return VerificationResult(False, "FAIL: Required outer border style differs",
                            details={"match_mode": "observed_outer_border_mismatch", "differences": differences})
                for edge in edges:
                    state = table["borders"].get(edge, {})
                    fields = {"line_style"}
                    if not outer_border_style_only and state.get("line_style") != 0:
                        fields |= {"line_width", "color"}
                    if physical_borders:
                        # Only a successful mixed observation permits refinement.
                        # Ordinary getter failures still fail this validation.
                        fields -= set(getattr(state, "mixed", ()))
                    error = require_observations({owner: state}, fields, context=f"table {index} border {edge}")
                    if error is not None:
                        return error
                if physical_borders:
                    error = validate_cell_borders(table, owner)
                    if error is not None:
                        return error
    return None
