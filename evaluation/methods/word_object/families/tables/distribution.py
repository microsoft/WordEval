"""Bounded geometry and insertion-inheritance checks for table distribution."""
from copy import deepcopy
import math
import zipfile
import xml.etree.ElementTree as ET

from evaluation.methods.base import VerificationResult


TOLERANCE_PT = 2.0
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def finite_dimensions(values, count):
    return (
        type(count) is int and isinstance(values, list) and len(values) == count and count > 0
        and all(isinstance(v, (int, float)) and not isinstance(v, bool)
                and math.isfinite(v) and 0 < v < 999999 for v in values)
    )


def row_distribution_check(gold_tables, pred_tables):
    """Do not accumulate a per-row tolerance into an arbitrary total drift."""
    if gold_tables and not pred_tables:
        return VerificationResult(False, "FAIL: Required table is missing")
    for role, tables in (("gold", gold_tables), ("prediction", pred_tables)):
        if not tables:
            return VerificationResult.error_result(f"{role}: table-height evidence is empty")
        for table in tables:
            heights = table.get("row_heights")
            if not finite_dimensions(heights, table.get("row_count", 0)):
                return VerificationResult.error_result(f"{role}: complete finite row heights required")
    differences = []
    for i, (gold, pred) in enumerate(zip(gold_tables, pred_tables)):
        gh, ph = gold["row_heights"], pred["row_heights"]
        if max(gh) - min(gh) > TOLERANCE_PT:
            return VerificationResult.error_result("Gold rows are not evenly distributed")
        if max(ph) - min(ph) > TOLERANCE_PT:
            differences.append({"table": i, "property": "row_height_spread", "prediction": max(ph)-min(ph)})
        if abs(sum(ph) - sum(gh)) > TOLERANCE_PT:
            differences.append({"table": i, "property": "total_row_height", "gold": sum(gh), "prediction": sum(ph)})
    if differences:
        return VerificationResult(False, "FAIL: Row distribution or total height differs", details={
            "differences": differences, "total_height_tolerance_pt": TOLERANCE_PT,
            "row_spread_tolerance_pt": TOLERANCE_PT})
    return None


def normalize_insert_split_distribution(input_tables, gold_tables, pred_tables, composition):
    """Recognize adjacent-row/column inheritance without erasing preservation.

    Word insertion APIs can inherit the neighboring row/column's dimensions.
    Distribution preserves each resulting scope's total, so compare those
    relations rather than the one default materialized in the reference.
    The caller still applies the ordinary table checks and saved preservation.
    """
    steps = composition.get("steps", [])
    expected = ["insert_row", "insert_column", "split_table", "distribute_columns", "distribute_rows"]
    if [s.get("operation") for s in steps] != expected:
        return VerificationResult.error_result("Unsupported insert/split/distribute step sequence")
    try:
        row_at = steps[0]["params"]["after_row"]
        col_at = steps[1]["params"]["after_col"]
        split_at = steps[2]["params"]["at_row"] - 1
        assert steps[3]["params"]["table_index"] == 1
        assert steps[4]["params"]["table_index"] == 2
        assert all(type(v) is int for v in (row_at, col_at, split_at))
        assert len(input_tables) == 1
        source = input_tables[0]
        nr, nc = source["row_count"], source["column_count"]
        assert 0 < row_at < nr and 0 < col_at < nc and split_at == row_at
        assert finite_dimensions(source["column_widths"], nc)
        assert finite_dimensions(source["row_heights"], nr)
        content = source["cell_contents"]
        assert len(content) == nr and all(len(row) == nc and all(isinstance(v, str) for v in row) for row in content)
    except (AssertionError, KeyError, TypeError):
        return VerificationResult.error_result("Input or insertion/distribution contract evidence is incomplete")

    expanded = deepcopy(content)
    expanded.insert(row_at, [""] * nc)
    for row in expanded:
        row.insert(col_at, "")
    expected_content = [expanded[:split_at], expanded[split_at:]]
    original_widths = source["column_widths"]
    original_heights = source["row_heights"]
    inherited_widths = original_widths[col_at-1:col_at+1]
    inherited_heights = original_heights[row_at-1:row_at+1]
    errors = []

    def close(a, b): return abs(a-b) <= TOLERANCE_PT

    def inspect(tables):
        issues = []
        if len(tables) != 2:
            return [{"property": "table_count", "expected": 2, "actual": len(tables)}]
        for index, (table, cells) in enumerate(zip(tables, expected_content)):
            observed = table.get("cell_contents")
            if not isinstance(observed, list) or any(not isinstance(row, list) or any(not isinstance(v, str) for v in row) for row in observed):
                raise ValueError("Complete cell-content evidence required")
            if (table.get("row_count"), table.get("column_count"), table.get("cell_contents")) != (len(cells), nc+1, cells):
                issues.append({"table": index, "property": "insert_split_content_positions"})
        if issues:
            return issues
        for table, cells in zip(tables, expected_content):
            if not finite_dimensions(table.get("column_widths"), nc+1) or not finite_dimensions(table.get("row_heights"), len(cells)):
                raise ValueError("Complete finite table geometry required")
        upper, lower = tables
        uw, lw = upper["column_widths"], lower["column_widths"]
        uh, lh = upper["row_heights"], lower["row_heights"]
        if not any(close(lw[col_at], w) for w in inherited_widths):
            issues.append({"property": "inserted_column_inheritance", "actual": lw[col_at], "allowed": inherited_widths})
        if any(not close(a,b) for a,b in zip(lw[:col_at]+lw[col_at+1:], original_widths)):
            issues.append({"property": "existing_lower_column_widths"})
        if max(uw)-min(uw) > TOLERANCE_PT or not close(sum(uw), sum(lw)):
            issues.append({"property": "upper_columns_distribution_and_total"})
        if any(not close(a,b) for a,b in zip(uh, original_heights[:split_at])):
            issues.append({"property": "upper_row_heights_preserved"})
        totals = [sum(original_heights[split_at:])+h for h in inherited_heights]
        if max(lh)-min(lh) > TOLERANCE_PT or not any(close(sum(lh), total) for total in totals):
            issues.append({"property": "lower_rows_distribution_and_total", "actual": lh, "allowed_totals": totals})
        return issues

    for role, tables in (("gold", gold_tables), ("prediction", pred_tables)):
        try:
            issues = inspect(tables)
        except (ValueError, TypeError, KeyError) as exc:
            return VerificationResult.error_result(f"{role}: {exc}")
        if role == "gold" and issues:
            return VerificationResult.error_result("Gold does not satisfy the input-bound table operations")
        if role == "prediction": errors = issues
    if errors:
        return VerificationResult(False, "FAIL: Insert/split/distribute semantics differ", details={"differences": errors})

    # Dimensions were validated against input and each other above. Only newly
    # inserted blank cells may inherit unrelated alignment/orientation defaults;
    # all pre-existing cells still receive the original strict comparisons.
    normalized = deepcopy(pred_tables)
    for ti, (gold, pred) in enumerate(zip(gold_tables, normalized)):
        pred["column_widths"] = list(gold["column_widths"])
        pred["row_heights"] = list(gold["row_heights"])
        for prop in ("cell_alignments", "cell_text_orientations", "cell_paragraph_alignments"):
            for role, table in (("gold", gold), ("prediction", pred)):
                matrix = table.get(prop)
                if not isinstance(matrix, list) or len(matrix) != len(expected_content[ti]) or any(not isinstance(row,list) or len(row)!=nc+1 or any(v is None for v in row) for row in matrix):
                    return VerificationResult.error_result(f"{role}: complete {prop} evidence required")
            for ri, row in enumerate(pred[prop]):
                for ci in range(nc+1):
                    if ci == col_at or (ti == 1 and ri == 0):
                        row[ci] = gold[prop][ri][ci]
    return normalized


def saved_insert_split_preservation(input_path, gold_path, pred_path, composition):
    """Preserve original cells while ignoring only insertion geometry metadata.

    The outer preservation policy excludes the target tables. Check their
    original cells here, including run formatting, fields and native objects.
    Newly inserted cells must contain only empty paragraphs and properties.
    """
    def tables(path):
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
        return [[list(row.findall(f"{{{W}}}tc")) for row in table.findall(f"{{{W}}}tr")]
                for table in root.findall(f"{{{W}}}body/{{{W}}}tbl")]

    def signature(node):
        if node.tag in {f"{{{W}}}tcW", f"{{{W}}}proofErr", f"{{{W}}}lang"}:
            return None
        attrs = tuple(sorted((k, v) for k, v in node.attrib.items()
                             if not k.rsplit("}", 1)[-1].startswith("rsid")
                             and k.rsplit("}", 1)[-1] not in {"paraId", "textId"}))
        return (node.tag, attrs, node.text,
                tuple(s for child in node if (s := signature(child)) is not None))

    def blank_cell(cell):
        # Property children are formatting, whereas fields, drawings, math,
        # hyperlinks, nested tables and nonempty text are inserted content.
        def blank(node):
            if node.tag in {f"{{{W}}}tcPr", f"{{{W}}}pPr", f"{{{W}}}rPr", f"{{{W}}}proofErr"}:
                return True
            if node.tag not in {f"{{{W}}}tc", f"{{{W}}}p", f"{{{W}}}r", f"{{{W}}}t"}:
                return False
            text = node.text if node.tag == f"{{{W}}}t" else (node.text or "").strip()
            return not text and all(blank(child) for child in node)
        return blank(cell)

    try:
        row_at = composition["steps"][0]["params"]["after_row"]
        col_at = composition["steps"][1]["params"]["after_col"]
        source = tables(input_path)
        if len(source) != 1 or not source[0] or not source[0][0]:
            raise ValueError("one rectangular source table required")
        rows, columns = len(source[0]), len(source[0][0])
        if any(len(row) != columns for row in source[0]):
            raise ValueError("rectangular source cells required")
        source = source[0]
        for role, path in (("gold", gold_path), ("prediction", pred_path)):
            target = tables(path)
            expected_rows = (row_at, rows + 1 - row_at)
            if len(target) != 2 or any(len(t) != n or any(len(r) != columns+1 for r in t)
                                       for t, n in zip(target, expected_rows)):
                raise ValueError(f"{role}: saved table structure disagrees with native evidence")
            differences = []
            for r in range(rows):
                ti, ri = (0, r) if r < row_at else (1, r + 1 - row_at)
                for c in range(columns):
                    ci = c if c < col_at else c+1
                    if signature(source[r][c]) != signature(target[ti][ri][ci]):
                        differences.append({"property": "original_cell_preservation", "row": r, "column": c})
            for ti, table in enumerate(target):
                for ri, row in enumerate(table):
                    for ci, cell in enumerate(row):
                        if (ci == col_at or (ti == 1 and ri == 0)) and not blank_cell(cell):
                            differences.append({"property": "inserted_cell_not_blank", "table": ti, "row": ri, "column": ci})
            if differences:
                if role == "gold":
                    return VerificationResult.error_result("Gold changes original cells or inserts nonblank cells")
                return VerificationResult(False, "FAIL: Saved cell preservation differs", details={"differences": differences})
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError, IndexError, TypeError, ValueError) as exc:
        return VerificationResult.error_result(f"Saved insertion preservation evidence unavailable: {exc}")
    return VerificationResult(True, "PASS: Original cells preserved and inserted cells blank")


def attach_operation_details(result, composition, *, evidence=None):
    """Keep the established intent schema without inventing raw-size failures.

    Successful input-bound geometry checks prove all five requested operations.
    When the complete relation fails, individual steps are not independently
    established: retain their identities and mark them unknown, while the
    primary result retains its specific failure evidence.
    """
    if evidence is not None:
        result.details["insertion_distribution"] = evidence
    intents = []
    flat = []
    supported_targets = {
        "insert_row": {("row_count", None)},
        "insert_column": {("column_count", None)},
        "split_table": {("table_count", None), ("row_count", None)},
        "distribute_columns": {("column_widths", 1)},
        "distribute_rows": {("row_heights", 2)},
    }
    for intent in composition.get("atomic_intents", []):
        targets = []
        for target in intent.get("verification_targets", []):
            index = target.get("table_index")
            supported = (target.get("extractor") == "table_properties"
                         and (index is None or type(index) is int)
                         and (target.get("property"), index) in supported_targets.get(intent.get("operation"), set()))
            passed = True if evidence is not None and supported else None
            item = {"target": deepcopy(target), "passed": passed,
                    "match_mode": "input_bound_operation_semantics" if passed else "not_independently_established"}
            if passed:
                for role, key in (("gold", "gold"), ("prediction", "pred")):
                    tables = evidence[role]
                    prop = target.get("property")
                    if prop == "table_count":
                        value = len(tables)
                    elif target.get("table_index"):
                        value = tables[target["table_index"]-1].get(prop)
                    else:
                        value = [t.get(prop) for t in tables]
                    item[key] = deepcopy(value)
            else:
                item["message"] = "This target was not independently established by the input-bound checks; see the primary result"
            targets.append(item)
            flat.append({"intent_id": intent["id"], "operation": intent.get("operation"),
                         "atomic_case": intent.get("atomic_case"), **item})
        intents.append({"intent_id": intent["id"], "operation": intent.get("operation"),
                        "atomic_case": intent.get("atomic_case"),
                        "passed": True if targets and all(t["passed"] is True for t in targets) else None,
                        "target_results": targets})
    if intents:
        result.details.update(atomic_results=intents, target_results=flat, atomic_summary={
            "passed_intents": sum(i["passed"] is True for i in intents), "total_intents": len(intents),
            "passed_targets": sum(t["passed"] is True for t in flat), "total_targets": len(flat),
            "unknown_intents": sum(i["passed"] is None for i in intents),
            "unknown_targets": sum(t["passed"] is None for t in flat)})
    return result
