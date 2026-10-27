"""families / tables / verify with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.core import DocumentOpenError
from evaluation.methods.word_object.families.tables import compare as families_tables_compare
from evaluation.methods.word_object.families.tables import distribution as table_distribution
from evaluation.methods.word_object.families.tables import (
    extract_saved as families_tables_extract_saved,
)
from pathlib import Path
from typing import Dict, List, Optional
from .observations import validate_tables, TABLE_FIELDS
from .geometry import physical_geometry_differences
from .borders import has_mixed_borders, cell_border_signature


def verify_table(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    com_primary_with_saved_guardrail,
    read_table_properties,
    doc_in_path: Optional[Path] = None,
) -> VerificationResult:
    """
    Verify table properties match between pred and gold.

    General verification: compares all table properties from gold to pred.
    Any difference is considered a failure.
    """
    idMso = spec.get("idMso", "") if spec else ""
    if idMso == "TableStylesGalleryWord":
        try:
            gold_tables = read_table_properties(str(gold_path))
            pred_tables = read_table_properties(str(pred_path))
        except DocumentOpenError:
            raise
        except Exception as com_error:
            gold_styles = families_tables_extract_saved.extract_saved_table_style_signatures(
                gold_path
            )
            pred_styles = families_tables_extract_saved.extract_saved_table_style_signatures(
                pred_path
            )
            return tag_verification_source(
                VerificationResult(
                    passed=gold_styles == pred_styles,
                    message=(
                        "PASS: Saved table style signature matches"
                        if gold_styles == pred_styles
                        else "FAIL: Saved table style signature differs"
                    ),
                    details={
                        "match_mode": "saved_ooxml_table_style",
                        "gold": gold_styles,
                        "pred": pred_styles,
                    },
                ),
                "saved_ooxml_fallback",
                word_com_error=str(com_error),
            )
        error = validate_tables(gold_tables, pred_tables, {"style"})
        if error is not None:
            return error
        com_differences = []
        if len(gold_tables) != len(pred_tables):
            com_differences.append(
                {
                    "property": "table_count",
                    "gold": len(gold_tables),
                    "pred": len(pred_tables),
                }
            )
        for index, (gold, pred) in enumerate(zip(gold_tables, pred_tables)):
            if gold.get("style") != pred.get("style"):
                com_differences.append(
                    {
                        "table_index": index,
                        "property": "style",
                        "gold": gold.get("style"),
                        "pred": pred.get("style"),
                    }
                )
        com_result = VerificationResult(
            passed=not com_differences,
            message=(
                "PASS: Word Object Model table style matches"
                if not com_differences
                else "FAIL: Word Object Model table style differs"
            ),
            details={"differences": com_differences},
        )
        if not com_result.passed:
            return tag_verification_source(com_result, "word_com")
        gold_styles = families_tables_extract_saved.extract_saved_table_style_signatures(gold_path)
        pred_styles = families_tables_extract_saved.extract_saved_table_style_signatures(pred_path)
        saved_result = VerificationResult(
            passed=gold_styles == pred_styles,
            message=(
                "PASS: Saved table style signature matches"
                if gold_styles == pred_styles
                else "FAIL: Saved table style signature differs"
            ),
            details={
                "match_mode": "saved_ooxml_table_style",
                "gold": gold_styles,
                "pred": pred_styles,
            },
        )
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_result)

    if idMso == "TableColumnsDelete":
        try:
            gold_com = read_table_properties(str(gold_path))
            pred_com = read_table_properties(str(pred_path))
        except DocumentOpenError:
            raise
        except Exception as com_error:
            return tag_verification_source(
                verify_saved_table_column_delete(pred_path, gold_path),
                "saved_ooxml_fallback",
                word_com_error=str(com_error),
            )
        error = validate_tables(gold_com, pred_com, {"row_count", "column_count", "cell_contents"})
        if error is not None:
            return error
        com_differences = []
        if len(gold_com) != len(pred_com):
            com_differences.append(
                {
                    "property": "table_count",
                    "gold": len(gold_com),
                    "pred": len(pred_com),
                }
            )
        for index, (gold, pred) in enumerate(zip(gold_com, pred_com)):
            for prop in ("row_count", "column_count", "cell_contents"):
                if gold.get(prop) != pred.get(prop):
                    com_differences.append(
                        {
                            "table_index": index,
                            "property": prop,
                            "gold": gold.get(prop),
                            "pred": pred.get(prop),
                        }
                    )
        com_result = VerificationResult(
            passed=not com_differences,
            message=(
                "PASS: Word Object Model deleted-column structure matches"
                if not com_differences
                else "FAIL: Word Object Model deleted-column structure differs"
            ),
            details={"differences": com_differences},
        )
        if not com_result.passed:
            return tag_verification_source(com_result, "word_com")
        saved_result = verify_saved_table_column_delete(pred_path, gold_path)
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_result)

    composition_commands = set(composition_command_ids(spec))
    saved_table_bundle = {
        "TableAutoFitWindow",
        "TableColumnsDistribute",
        "TableCellAlignMiddleCenter",
    }
    if composition_commands == saved_table_bundle:
        return com_primary_with_saved_guardrail(
            lambda: verify_table_window_distribute_align_com(
                pred_path, gold_path, read_table_properties=read_table_properties
            ),
            lambda: verify_saved_table_window_distribute_align(pred_path, gold_path),
        )
    if composition_commands == {
        "TableStylesGalleryWord",
        "TextDirectionGallery",
        "BorderStylesGallery",
    }:
        gold_tables = read_table_properties(str(gold_path))
        pred_tables = read_table_properties(str(pred_path))
        error = validate_tables(gold_tables, pred_tables, {"row_count", "column_count", "cell_contents", "style", "cell_text_orientations", "borders"}, outer_border_style_only=True)
        if error is not None:
            return error
        return families_tables_compare.verify_additive_table_style_direction_border(
            gold_tables, pred_tables
        )

    gold_tables = read_table_properties(str(gold_path))
    pred_tables = read_table_properties(str(pred_path))

    required = set(TABLE_FIELDS)
    if idMso == "TableDelete":
        required = set()
    elif idMso in {"TableAutoFitContents", "TableAutoFitWindow", "TableAutoFitFixedColumnWidth"}:
        required = {"row_count", "column_count", "cell_contents"}
    if idMso in {"TableColumnsInsertLeft", "TableColumnsInsertRight"}:
        required = {"row_count", "column_count", "cell_contents", "column_widths"}
    interaction = (spec or {}).get("composition", {}).get("interaction_class", "")
    if interaction == "insert_then_border":
        required = {"row_count", "column_count", "cell_contents", "borders"}
    elif interaction == "insert_then_align":
        required = {"row_count", "column_count", "cell_contents", "cell_alignments", "cell_paragraph_alignments"}
    elif interaction == "insert_then_text_align":
        required = {"row_count", "column_count", "cell_contents", "cell_paragraph_alignments"}
    physical_geometry = (
        required == TABLE_FIELDS
        and idMso not in {"TableRowsDistribute", "TableColumnsDistribute"}
        and interaction != "insert_row_column_split_then_distribute_scopes"
        and any(table.get(field) is None for table in (gold_tables or []) + (pred_tables or [])
                for field in ("column_widths", "row_heights"))
    )
    if physical_geometry:
        required = required - {"column_widths", "row_heights"} | {"cell_geometry"}
    physical_borders = (
        "borders" in required and interaction != "insert_then_border"
        and any(has_mixed_borders(table) for table in (gold_tables or []) + (pred_tables or []))
    )
    if physical_borders:
        for owner, path, tables in (("gold", gold_path, gold_tables), ("pred", pred_path, pred_tables)):
            if any("cell_borders" not in table for table in tables):
                refined = read_table_properties(str(path), include_cell_borders=True)
                if len(refined) != len(tables):
                    return VerificationResult.error_result("Table collection changed during border refinement", observation_owner=owner)
                for table, item in zip(tables, refined):
                    # Keep the original observations and their failures. Only
                    # add the finer evidence that was not requested initially.
                    table["cell_borders"] = item.get("cell_borders")
                    if "cell_borders" in getattr(item, "unavailable", {}) and hasattr(table, "record_unavailable"):
                        table.record_unavailable("cell_borders", item.unavailable["cell_borders"])
    if interaction == "insert_row_column_split_then_distribute_scopes":
        # The input-bound operation check validates the resulting table scopes
        # before asking for dimensions within them. A wrong split can produce
        # irregular columns; their unreadability must not hide the observed
        # wrong split. Complete preservation is checked on the normalized
        # projection only after these operation prerequisites succeed.
        error = validate_tables(gold_tables, [], required)
        if error is None:
            error = validate_tables(gold_tables, pred_tables, {"row_count", "column_count", "cell_contents"})
    else:
        error = validate_tables(gold_tables, pred_tables, required, allow_empty=idMso == "TableDelete",
                                outer_border_style_only=interaction == "insert_then_border", physical_borders=physical_borders)
    if error is not None:
        return error

    interaction_class = (spec or {}).get("composition", {}).get("interaction_class", "")

    if interaction_class == "insert_row_column_split_then_distribute_scopes":
        if doc_in_path is None:
            return VerificationResult.error_result("Input document required for insertion/distribution preservation")
        input_tables = read_table_properties(str(doc_in_path))
        error = validate_tables(input_tables, [], TABLE_FIELDS)
        if error is not None:
            return error
        normalized = table_distribution.normalize_insert_split_distribution(
            input_tables, gold_tables, pred_tables, spec["composition"]
        )
        if isinstance(normalized, VerificationResult):
            return table_distribution.attach_operation_details(normalized, spec["composition"])
        evidence = {
            "match_mode": "input_bound_neighbor_inheritance",
            "input": input_tables, "gold": gold_tables, "prediction": pred_tables,
            "tolerance_pt": table_distribution.TOLERANCE_PT,
            "operation_checks": [
                {"operation": step["operation"], "passed": True,
                 "match_mode": "input_bound_operation_semantics"}
                for step in spec["composition"]["steps"]
            ],
        }
        saved = table_distribution.saved_insert_split_preservation(
            doc_in_path, gold_path, pred_path, spec["composition"]
        )
        if not saved.passed:
            return table_distribution.attach_operation_details(saved, spec["composition"], evidence=evidence)
        result = verify_table(
            pred_path, gold_path, {},
            com_primary_with_saved_guardrail=com_primary_with_saved_guardrail,
            read_table_properties=lambda path: gold_tables if path == str(gold_path) else normalized,
        )
        # Keep public observations native; only the internal comparison uses
        # the validated dimension/default projection.
        result.details.update(idMso=idMso, gold_tables=gold_tables, pred_tables=pred_tables,
                              comparison_mode="input_bound_insert_split_distribution")
        return table_distribution.attach_operation_details(result, spec["composition"], evidence=evidence)

    if idMso == "TableRowsDistribute":
        geometry_failure = table_distribution.row_distribution_check(gold_tables, pred_tables)
        if geometry_failure is not None:
            return geometry_failure

    if idMso in {"TableColumnsInsertLeft", "TableColumnsInsertRight"}:
        return families_tables_compare.verify_table_column_insert_semantics(
            gold_tables, pred_tables, idMso
        )

    if interaction_class == "insert_then_border":
        return families_tables_compare.verify_insert_then_border_semantics(gold_tables, pred_tables)
    if interaction_class == "insert_then_align":
        return families_tables_compare.verify_inserted_table_properties(
            gold_tables,
            pred_tables,
            (
                "row_count",
                "column_count",
                "cell_contents",
                "cell_alignments",
                "cell_paragraph_alignments",
            ),
            "cell alignment",
        )
    if interaction_class == "insert_then_text_align":
        return families_tables_compare.verify_inserted_table_properties(
            gold_tables,
            pred_tables,
            ("row_count", "column_count", "cell_contents", "cell_paragraph_alignments"),
            "cell text alignment",
        )

    # Special case: TableDelete should result in no tables
    if idMso == "TableDelete":
        if len(gold_tables) == 0 and len(pred_tables) == 0:
            return VerificationResult(
                passed=True,
                message="PASS: Table deleted (no tables in either document)",
                details={"idMso": idMso},
            )
        elif len(pred_tables) > 0:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Table not deleted (pred has {len(pred_tables)} table(s))",
                details={"idMso": idMso, "pred_table_count": len(pred_tables)},
            )

    # Special case: AutoFit commands — verify AutoFit semantic flags
    # (AllowAutoFit + PreferredWidthType) rather than raw column widths.
    # Raw widths return 9999999 (wdUndefined) in AutoFit dynamic state, so
    # they can serialize inconsistently between gold and pred even when
    # semantically equivalent.
    if idMso in {"TableAutoFitContents", "TableAutoFitWindow", "TableAutoFitFixedColumnWidth"}:
        return verify_table_autofit(gold_tables, pred_tables, pred_path, gold_path, idMso)

    # Check table count
    if len(gold_tables) != len(pred_tables):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Table count mismatch (gold={len(gold_tables)}, pred={len(pred_tables)})",
            details={
                "gold_table_count": len(gold_tables),
                "pred_table_count": len(pred_tables),
                "idMso": idMso,
            },
        )

    if not gold_tables:
        return VerificationResult(
            passed=False, message="FAIL: No tables found in gold document", details={"idMso": idMso}
        )

    differences = []

    for i, (gold_table, pred_table) in enumerate(zip(gold_tables, pred_tables)):
        table_diffs = []

        # Compare row count
        if gold_table.get("row_count") != pred_table.get("row_count"):
            table_diffs.append(
                {
                    "property": "row_count",
                    "gold": gold_table.get("row_count"),
                    "pred": pred_table.get("row_count"),
                }
            )

        # Compare column count
        if gold_table.get("column_count") != pred_table.get("column_count"):
            table_diffs.append(
                {
                    "property": "column_count",
                    "gold": gold_table.get("column_count"),
                    "pred": pred_table.get("column_count"),
                }
            )

        # Compare cell contents
        if gold_table.get("cell_contents") != pred_table.get("cell_contents"):
            table_diffs.append(
                {
                    "property": "cell_contents",
                    "gold": gold_table.get("cell_contents"),
                    "pred": pred_table.get("cell_contents"),
                }
            )

        # Compare cell alignments
        if gold_table.get("cell_alignments") != pred_table.get("cell_alignments"):
            table_diffs.append(
                {
                    "property": "cell_alignments",
                    "gold": gold_table.get("cell_alignments"),
                    "pred": pred_table.get("cell_alignments"),
                }
            )

        # Compare cell text orientations
        if gold_table.get("cell_text_orientations") != pred_table.get("cell_text_orientations"):
            table_diffs.append(
                {
                    "property": "cell_text_orientations",
                    "gold": gold_table.get("cell_text_orientations"),
                    "pred": pred_table.get("cell_text_orientations"),
                }
            )

        # Compare cell paragraph alignments (horizontal text alignment)
        if gold_table.get("cell_paragraph_alignments") != pred_table.get(
            "cell_paragraph_alignments"
        ):
            table_diffs.append(
                {
                    "property": "cell_paragraph_alignments",
                    "gold": gold_table.get("cell_paragraph_alignments"),
                    "pred": pred_table.get("cell_paragraph_alignments"),
                }
            )

        # Compare table style
        if gold_table.get("style") != pred_table.get("style"):
            table_diffs.append(
                {
                    "property": "style",
                    "gold": gold_table.get("style"),
                    "pred": pred_table.get("style"),
                }
            )

        # Compare borders
        gold_borders = gold_table.get("borders", {})
        pred_borders = pred_table.get("borders", {})
        if physical_borders:
            gold_cells, pred_cells = cell_border_signature(gold_table), cell_border_signature(pred_table)
            if gold_cells != pred_cells:
                table_diffs.append({"property": "physical_cell_borders", "gold": gold_cells, "pred": pred_cells})
        for border_name in ([] if physical_borders else ["top", "left", "bottom", "right", "horizontal", "vertical"]):
            gold_border = gold_borders.get(border_name, {})
            pred_border = pred_borders.get(border_name, {})
            for prop in ["line_style", "line_width", "color"]:
                gold_val = gold_border.get(prop)
                pred_val = pred_border.get(prop)
                if gold_val != pred_val:
                    table_diffs.append(
                        {
                            "property": f"border_{border_name}_{prop}",
                            "gold": gold_val,
                            "pred": pred_val,
                        }
                    )

        if physical_geometry:
            table_diffs.extend(physical_geometry_differences(
                gold_table["cell_geometry"], pred_table["cell_geometry"]))
        else:
            # Compare column widths (with tolerance for floating point)
            gold_widths = gold_table.get("column_widths", [])
            pred_widths = pred_table.get("column_widths", [])
            if len(gold_widths) == len(pred_widths):
                tolerance = 2.0  # 2 points tolerance
                for j, (gw, pw) in enumerate(zip(gold_widths, pred_widths)):
                    if abs(gw - pw) > tolerance:
                        table_diffs.append({"property": f"column_width_{j}", "gold": gw, "pred": pw})
            elif gold_widths or pred_widths:
                table_diffs.append(
                    {"property": "column_widths", "gold": gold_widths, "pred": pred_widths}
                )

            # Compare row heights (with tolerance for floating point)
            gold_heights = gold_table.get("row_heights", [])
            pred_heights = pred_table.get("row_heights", [])
            if len(gold_heights) == len(pred_heights):
                tolerance = 2.0  # 2 points tolerance
                for j, (gh, ph) in enumerate(zip(gold_heights, pred_heights)):
                    if abs(gh - ph) > tolerance:
                        table_diffs.append({"property": f"row_height_{j}", "gold": gh, "pred": ph})
            elif gold_heights or pred_heights:
                table_diffs.append(
                    {"property": "row_heights", "gold": gold_heights, "pred": pred_heights}
                )

        if table_diffs:
            differences.append({"table_index": i, "differences": table_diffs})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} table(s) have differences",
            details={
                "differences": differences,
                "idMso": idMso,
                "gold_tables": gold_tables,
                "pred_tables": pred_tables,
            },
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: All {len(gold_tables)} table(s) match",
        details={
            "table_count": len(gold_tables),
            "idMso": idMso,
            "gold_structure": (
                {
                    "rows": gold_tables[0].get("row_count"),
                    "cols": gold_tables[0].get("column_count"),
                }
                if gold_tables
                else {}
            ),
        },
    )


def verify_saved_table_column_delete(pred_path: Path, gold_path: Path) -> VerificationResult:
    gold_tables = families_tables_extract_saved.extract_saved_table_composition_signatures(
        gold_path
    )
    pred_tables = families_tables_extract_saved.extract_saved_table_composition_signatures(
        pred_path
    )
    differences = []
    if len(gold_tables) != len(pred_tables):
        differences.append(
            {
                "property": "table_count",
                "gold": len(gold_tables),
                "pred": len(pred_tables),
            }
        )
    for index, (gold, pred) in enumerate(zip(gold_tables, pred_tables)):
        for prop in ("row_count", "column_count", "cell_contents"):
            if gold.get(prop) != pred.get(prop):
                differences.append(
                    {
                        "table_index": index,
                        "property": prop,
                        "gold": gold.get(prop),
                        "pred": pred.get(prop),
                    }
                )
    return VerificationResult(
        passed=not differences,
        message=(
            "PASS: Saved deleted-column structure matches"
            if not differences
            else "FAIL: Saved deleted-column structure differs"
        ),
        details={
            "match_mode": "saved_ooxml_table_column_delete",
            "differences": differences,
        },
    )


def verify_table_window_distribute_align_com(
    pred_path: Path, gold_path: Path, *, read_table_properties
) -> VerificationResult:
    """Check rendered table width/distribution/alignment through WOM."""
    gold_tables = read_table_properties(str(gold_path))
    pred_tables = read_table_properties(str(pred_path))
    error = validate_tables(gold_tables, pred_tables, {
        "row_count", "column_count", "cell_contents", "cell_alignments",
        "cell_paragraph_alignments", "column_widths"})
    if error is not None:
        return error
    differences = []
    if len(gold_tables) != len(pred_tables):
        differences.append(
            {
                "property": "table_count",
                "gold": len(gold_tables),
                "pred": len(pred_tables),
            }
        )
    for index, (gold, pred) in enumerate(zip(gold_tables, pred_tables)):
        for prop in (
            "row_count",
            "column_count",
            "cell_contents",
            "cell_alignments",
            "cell_paragraph_alignments",
        ):
            if gold.get(prop) != pred.get(prop):
                differences.append(
                    {
                        "table_index": index,
                        "property": prop,
                        "gold": gold.get(prop),
                        "pred": pred.get(prop),
                    }
                )
        gold_widths = gold.get("column_widths") or []
        pred_widths = pred.get("column_widths") or []
        tolerances = families_tables_compare.rendered_table_width_tolerances(
            gold_widths, pred_widths
        )
        if pred_widths and (max(pred_widths) - min(pred_widths) > tolerances["column_spread"]):
            differences.append(
                {
                    "table_index": index,
                    "property": "rendered_column_width_spread",
                    "gold": gold_widths,
                    "pred": pred_widths,
                    "tolerance": tolerances["column_spread"],
                }
            )
        if (
            gold_widths
            and pred_widths
            and abs(sum(gold_widths) - sum(pred_widths)) > tolerances["total_width"]
        ):
            differences.append(
                {
                    "table_index": index,
                    "property": "rendered_total_width",
                    "gold": sum(gold_widths),
                    "pred": sum(pred_widths),
                    "tolerance": tolerances["total_width"],
                }
            )
    return VerificationResult(
        passed=not differences,
        message=(
            "PASS: Word Object Model table window/distribution/alignment matches"
            if not differences
            else "FAIL: Word Object Model table window/distribution/alignment differs"
        ),
        details={"differences": differences},
    )


def verify_saved_table_window_distribute_align(
    pred_path: Path, gold_path: Path
) -> VerificationResult:
    """Verify the AutoFit-window + distribute-columns + align bundle."""
    gold = families_tables_extract_saved.extract_saved_table_composition_signatures(gold_path)
    pred = families_tables_extract_saved.extract_saved_table_composition_signatures(pred_path)
    differences = []
    if len(gold) != len(pred):
        differences.append({"property": "table_count", "gold": len(gold), "pred": len(pred)})
    for index, (gold_table, pred_table) in enumerate(zip(gold, pred)):
        for prop in ("row_count", "column_count", "cell_contents"):
            if gold_table.get(prop) != pred_table.get(prop):
                differences.append(
                    {
                        "table_index": index,
                        "property": prop,
                        "gold": gold_table.get(prop),
                        "pred": pred_table.get(prop),
                    }
                )
        gold_spread = families_tables_compare.preferred_column_spread(gold_table)
        pred_spread = families_tables_compare.preferred_column_spread(pred_table)
        # pct widths use fiftieths of a percent; dxa widths use twips.
        pred_types = {
            kind
            for column in pred_table.get("preferred_columns", [])
            for kind, value in column
            if value is not None
        }
        preferred_values = [
            float(value)
            for column in pred_table.get("preferred_columns", [])
            for kind, value in column
            if value is not None
        ]
        preferred_average = (
            sum(preferred_values) / len(preferred_values) if preferred_values else 0.0
        )
        tolerance = max(2.0, 0.003 * preferred_average)
        if pred_spread is None or pred_spread > tolerance:
            differences.append(
                {
                    "table_index": index,
                    "property": "distributed_column_preferred_width_spread",
                    "gold": gold_spread,
                    "pred": pred_spread,
                    "tolerance": tolerance,
                }
            )
        if gold_table.get("cell_alignments") != pred_table.get("cell_alignments"):
            differences.append(
                {
                    "table_index": index,
                    "property": "cell_alignments",
                    "gold": gold_table.get("cell_alignments"),
                    "pred": pred_table.get("cell_alignments"),
                }
            )

    gold_flags = families_tables_extract_saved.extract_table_autofit_flags(gold_path)
    pred_flags = families_tables_extract_saved.extract_table_autofit_flags(pred_path)
    if len(gold_flags) != len(pred_flags):
        differences.append(
            {
                "property": "autofit_table_count",
                "gold": len(gold_flags),
                "pred": len(pred_flags),
            }
        )
    for index, pred_flag in enumerate(pred_flags):
        if not (
            pred_flag.get("allow_autofit")
            and pred_flag.get("preferred_width_type") == 2
            and pred_flag.get("preferred_width") is not None
            and abs(pred_flag["preferred_width"] - 100.0) <= 0.1
        ):
            differences.append(
                {
                    "table_index": index,
                    "property": "autofit_window",
                    "gold": gold_flags[index] if index < len(gold_flags) else None,
                    "pred": pred_flag,
                }
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Saved table window/distribution/alignment bundle "
                f"differs ({len(differences)} difference(s))"
            ),
            details={
                "differences": differences,
                "verification_source": "saved_ooxml",
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Saved table window/distribution/alignment bundle matches",
        details={"verification_source": "saved_ooxml"},
    )


def verify_table_autofit(
    gold_tables: List[Dict], pred_tables: List[Dict], pred_path: Path, gold_path: Path, idMso: str
) -> VerificationResult:
    """
    Verify TableAutoFit* commands using AutoFit semantic flags.

    Dispatched from `_verify_table` for:
      - TableAutoFitContents        (PreferredWidthType = wdPreferredWidthAuto)
      - TableAutoFitWindow          (PreferredWidthType = wdPreferredWidthPercent)
      - TableAutoFitFixedColumnWidth (PreferredWidthType = wdPreferredWidthPoints)

    Why this exists: raw column Width values return 9999999 (wdUndefined)
    in AutoFit dynamic state. Comparing them across save/load cycles is
    unreliable — pred and gold may differ numerically while being
    semantically identical AutoFit tables.  This method compares the
    AutoFit flags directly and verifies that table structure/content was
    preserved (AutoFit should not modify rows, columns, or cell contents).

    Structural preservation is cross-checked against the extracted table
    properties that were already computed in _verify_table; AutoFit flags
    themselves are queried from Word COM here (not present in the generic
    extraction).
    """
    diffs = []

    if len(gold_tables) != len(pred_tables):
        return VerificationResult(
            passed=False,
            message=(
                f"FAIL: Table count mismatch " f"(gold={len(gold_tables)}, pred={len(pred_tables)})"
            ),
            details={
                "gold_table_count": len(gold_tables),
                "pred_table_count": len(pred_tables),
                "idMso": idMso,
            },
        )

    # --- Structural preservation: rows, cols, cell contents ---
    for i, (g, p) in enumerate(zip(gold_tables, pred_tables)):
        for prop in ("row_count", "column_count", "cell_contents", "style"):
            if g.get(prop) != p.get(prop):
                diffs.append(
                    {
                        "table_index": i,
                        "property": prop,
                        "gold": g.get(prop),
                        "pred": p.get(prop),
                    }
                )

    try:
        gold_flags = families_tables_extract_saved.extract_table_autofit_flags(gold_path)
        pred_flags = families_tables_extract_saved.extract_table_autofit_flags(pred_path)
    except DocumentOpenError:
        raise
    except Exception as e:
        return VerificationResult(
            passed=False,
            message=f"ERROR: AutoFit flag extraction failed: {e}",
            error=f"AutoFit flag extraction failed: {e}",
            details={"idMso": idMso, "error": str(e)},
        )

    # --- AutoFit flag comparison (PreferredWidthType is the primary discriminator) ---
    for i, (g, p) in enumerate(zip(gold_flags, pred_flags)):
        if g["allow_autofit"] != p["allow_autofit"]:
            diffs.append(
                {
                    "table_index": i,
                    "property": "allow_autofit",
                    "gold": g["allow_autofit"],
                    "pred": p["allow_autofit"],
                }
            )
        if g["preferred_width_type"] != p["preferred_width_type"]:
            diffs.append(
                {
                    "table_index": i,
                    "property": "preferred_width_type",
                    "gold": g["preferred_width_type"],
                    "pred": p["preferred_width_type"],
                }
            )
        gpw, ppw = g["preferred_width"], p["preferred_width"]
        if gpw is not None and ppw is not None:
            if abs(gpw - ppw) > 2.0:
                diffs.append(
                    {
                        "table_index": i,
                        "property": "preferred_width",
                        "gold": gpw,
                        "pred": ppw,
                    }
                )

    if diffs:
        return VerificationResult(
            passed=False,
            message=f"FAIL: AutoFit verification — {len(diffs)} difference(s)",
            details={"idMso": idMso, "differences": diffs},
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: AutoFit properties match ({idMso})",
        details={
            "idMso": idMso,
            "table_count": len(gold_tables),
            "preferred_width_type": gold_flags[0]["preferred_width_type"] if gold_flags else None,
        },
    )
