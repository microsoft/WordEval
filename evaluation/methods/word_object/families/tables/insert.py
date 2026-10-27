"""families / tables / insert with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.tables import (
    insert_saved as families_tables_insert_saved,
)
from pathlib import Path
from typing import Dict, Optional
from .observations import validate_tables
from ...preservation.context import verify_saved_insertion_context


def verify_tables(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    com_primary_with_saved_guardrail,
    error_with_details,
    read_table_properties,
) -> VerificationResult:
    if (spec or {}).get("idMso") == "TableInsertGallery":
        return com_primary_with_saved_guardrail(
            lambda: verify_tables_com(
                pred_path, gold_path, spec, read_table_properties=read_table_properties
            ),
            lambda: verify_table_insert_saved(
                pred_path, gold_path, error_with_details=error_with_details
            ),
            saved_pass_after_com_failure=True,
        )
    return verify_tables_com(
        pred_path, gold_path, spec, read_table_properties=read_table_properties
    )


def verify_table_insert_saved(
    pred_path: Path, gold_path: Path, *, error_with_details
) -> VerificationResult:
    gold = families_tables_insert_saved.extract_saved_table_insert_signature(gold_path)
    pred = families_tables_insert_saved.extract_saved_table_insert_signature(pred_path)
    if not gold.get("tables"):
        return error_with_details(
            "Gold saved inserted-table evidence is missing",
            {
                "match_mode": "saved_ooxml_table_insert_structure",
                "gold": gold,
            },
        )
    if pred != gold:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved inserted-table structure differs",
            details={
                "match_mode": "saved_ooxml_table_insert_structure",
                "gold": gold,
                "pred": pred,
            },
        )
    return VerificationResult(
        passed=True,
        message=(
            "PASS: Saved inserted-table structure and visible content "
            f"match ({len(gold['tables'])} table(s))"
        ),
        details={
            "match_mode": "saved_ooxml_table_insert_structure",
            "fallback_contract": "complete",
            "signature": gold,
        },
    )


def verify_tables_com(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_table_properties
) -> VerificationResult:
    """
    Verify table structure/style/content matches gold.

    For QuickTables gallery insertions, exact OpenXML can vary while Word-level
    table structure and rendered content/style are the meaningful signals.
    """
    gold_tables = read_table_properties(str(gold_path))
    pred_tables = read_table_properties(str(pred_path))

    required = {"row_count", "column_count", "cell_contents"}
    if (spec or {}).get("idMso") == "QuickTablesInsertGallery":
        required.add("style")
    error = validate_tables(gold_tables, pred_tables, required)
    if error is not None:
        return error

    if len(gold_tables) != len(pred_tables):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Table count mismatch (pred={len(pred_tables)}, gold={len(gold_tables)})",
            details={"pred_count": len(pred_tables), "gold_count": len(gold_tables)},
        )

    def _strip_cell(value):
        """Strip trailing Word control chars (\\r\\x07) from cell text."""
        if value is None:
            return None
        return str(value).strip("\r\x07")

    differences = []
    for i, (gt, pt) in enumerate(zip(gold_tables, pred_tables), start=1):
        table_diffs = []

        if gt.get("row_count") != pt.get("row_count") or gt.get("column_count") != pt.get(
            "column_count"
        ):
            table_diffs.append(
                {
                    "property": "dimensions",
                    "gold": {"rows": gt.get("row_count"), "cols": gt.get("column_count")},
                    "pred": {"rows": pt.get("row_count"), "cols": pt.get("column_count")},
                }
            )

        if (spec or {}).get("idMso") == "QuickTablesInsertGallery" and (
            gt.get("style") or None
        ) != (pt.get("style") or None):
            table_diffs.append(
                {
                    "property": "style",
                    "gold": gt.get("style"),
                    "pred": pt.get("style"),
                }
            )

        gold_cells = [[_strip_cell(c) for c in row] for row in gt.get("cell_contents", [])]
        pred_cells = [[_strip_cell(c) for c in row] for row in pt.get("cell_contents", [])]
        if gold_cells != pred_cells:
            table_diffs.append(
                {
                    "property": "cell_contents",
                    "gold": gold_cells,
                    "pred": pred_cells,
                    "match_mode": "exact",
                }
            )

        if table_diffs:
            differences.append(
                {
                    "table_index": i,
                    "differences": table_diffs,
                }
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} table(s) differ",
            details={"differences": differences},
        )

    if (spec or {}).get("idMso") in {"SortDialogClassic", "QuickTablesInsertGallery"}:
        preservation = verify_saved_insertion_context(pred_path, gold_path, target_kind="table")
        if preservation is not None:
            return preservation

    return VerificationResult(
        passed=True,
        message=f"PASS: Tables match ({len(gold_tables)} table(s))",
        details={"table_count": len(gold_tables)},
    )
