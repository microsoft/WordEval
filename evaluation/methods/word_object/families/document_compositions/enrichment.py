"""families / document_compositions / enrichment with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.families.document_compositions import (
    extract_saved as families_document_compositions_extract_saved,
)
from evaluation.methods.word_object.families.document_compositions import (
    preservation as families_document_compositions_preservation,
)
from evaluation.methods.word_object.families.headers import verify as families_headers_verify
from evaluation.methods.word_object.families.notes import verify as families_notes_verify
from evaluation.methods.word_object.families.references import (
    bibliography as families_references_bibliography,
)
from evaluation.methods.word_object.families.references import verify as families_references_verify
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, List, Optional


def verify_document_enrichment_saved(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
) -> VerificationResult:
    gold = families_document_compositions_extract_saved.extract_saved_enrichment_signature(
        gold_path
    )
    pred = families_document_compositions_extract_saved.extract_saved_enrichment_signature(
        pred_path
    )
    intents = (spec or {}).get("composition", {}).get("atomic_intents", [])
    requested_operations = [
        intent.get("operation") for intent in intents if intent.get("operation")
    ]

    def note_errors(owner: str, operation: str, value: Dict[str, Any]) -> List[str]:
        references = value.get("references", [])
        errors = []
        if len(references) != 1:
            errors.append(f"{owner} {operation} reference count is {len(references)}, not one")
        if value.get("relationship_count") != 1:
            errors.append(f"{owner} {operation} relationship count is not one")
        if any(not reference.get("valid_id") for reference in references):
            errors.append(f"{owner} {operation} reference ID is invalid")
        if any(not (reference.get("note_text") or "").strip() for reference in references):
            errors.append(f"{owner} {operation} text is empty")
        if value.get("orphan_notes"):
            errors.append(f"{owner} {operation} has orphan note definitions")
        return errors

    gold_errors = []
    pred_structure_errors = []
    for operation in ("footnote", "endnote"):
        if operation not in requested_operations:
            continue
        gold_errors.extend(note_errors("gold", operation, gold[operation]))
        pred_structure_errors.extend(note_errors("predicted", operation, pred[operation]))
    if "page_number_header" in requested_operations and gold["header_page_field_count"] < 1:
        gold_errors.append("gold header has no PAGE field")
    if "page_number_footer" in requested_operations and gold["footer_page_field_count"] < 1:
        gold_errors.append("gold footer has no PAGE field")
    if "citation" in requested_operations and not gold["citation_tags"]:
        gold_errors.append("gold has no CITATION field tag")
    if "bibliography" in requested_operations and gold["bibliography_field_count"] < 1:
        gold_errors.append("gold has no BIBLIOGRAPHY field")
    if gold_errors:
        return policies_content_operations.error_with_details(
            "Gold saved document enrichment evidence is missing or invalid",
            {
                "match_mode": "saved_ooxml_document_enrichment",
                "errors": gold_errors,
                "gold": gold,
            },
        )

    available_checks = {
        "footnote": (
            not any("footnote" in error for error in pred_structure_errors)
            and pred["footnote"] == gold["footnote"]
        ),
        "endnote": (
            not any("endnote" in error for error in pred_structure_errors)
            and pred["endnote"] == gold["endnote"]
        ),
        "page_number_header": (
            pred["header_page_field_count"] == gold["header_page_field_count"]
            and gold["header_page_field_count"] > 0
        ),
        "page_number_footer": (
            pred["footer_page_field_count"] == gold["footer_page_field_count"]
            and gold["footer_page_field_count"] > 0
        ),
        "citation": (
            bool(gold["citation_tags"]) and pred["citation_tags"] == gold["citation_tags"]
        ),
        "bibliography": (
            gold["bibliography_field_count"] > 0
            and pred["bibliography_field_count"] == gold["bibliography_field_count"]
        ),
    }
    value_keys = {
        "footnote": "footnote",
        "endnote": "endnote",
        "page_number_header": "header_page_field_count",
        "page_number_footer": "footer_page_field_count",
        "citation": "citation_tags",
        "bibliography": "bibliography_field_count",
    }
    checks = {
        operation: available_checks[operation]
        for operation in requested_operations
        if operation in available_checks
    }
    unsupported_operations = sorted(
        {operation for operation in requested_operations if operation not in available_checks}
    )
    atomic_results = []
    flat_targets = []
    for intent in intents:
        operation = intent.get("operation")
        supported = operation in checks
        passed = bool(checks.get(operation, False)) if supported else False
        key = value_keys.get(operation)
        target_results = []
        for target in intent.get("verification_targets", []) or []:
            target_result = {
                "target": target,
                "passed": passed,
                "gold": gold.get(key),
                "pred": pred.get(key),
            }
            if not supported:
                target_result["error"] = (
                    "Saved-state fallback does not verify operation " f"{operation}"
                )
            target_results.append(target_result)
            flat_targets.append(
                {
                    **target_result,
                    "intent_id": intent.get("id"),
                    "operation": operation,
                    "atomic_case": intent.get("atomic_case"),
                }
            )
        atomic_results.append(
            {
                "intent_id": intent.get("id"),
                "operation": operation,
                "atomic_case": intent.get("atomic_case"),
                "passed": passed and bool(target_results),
                "target_results": target_results,
            }
        )

    failures = [operation for operation, passed in checks.items() if not passed]
    fallback_error = None
    if not failures and unsupported_operations:
        fallback_error = "Saved-state fallback cannot verify requested operation(s): " + ", ".join(
            unsupported_operations
        )
    return VerificationResult(
        passed=not failures and not unsupported_operations,
        message=(
            f"ERROR: {fallback_error}"
            if fallback_error
            else (
                "PASS: Saved document enrichment semantics match"
                if not failures and not unsupported_operations
                else "FAIL: Saved document enrichment semantics differ " f"({', '.join(failures)})"
            )
        ),
        details={
            "match_mode": "saved_ooxml_document_enrichment",
            "failures": failures,
            "unsupported_operations": unsupported_operations,
            "pred_structure_errors": pred_structure_errors,
            "gold": gold,
            "pred": pred,
            "atomic_results": atomic_results,
            "target_results": flat_targets,
            "atomic_summary": {
                "passed_intents": sum(1 for result in atomic_results if result.get("passed")),
                "total_intents": len(atomic_results),
                "passed_targets": sum(1 for result in flat_targets if result.get("passed")),
                "total_targets": len(flat_targets),
            },
        },
        error=fallback_error,
    )


def verify_document_enrichment_composition(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    add_composition_atomic_details,
    doc_in_path,
    read_bibliography_info,
    read_document_text,
    read_dynamic_reference_state,
    read_footnote_endnote_content,
    read_header_footer_content,
    summarize_sub_result,
) -> VerificationResult:
    """Verify document enrichment: footnotes, endnotes, header/footer, citation, bibliography."""
    commands = composition_command_ids(spec)
    if gold_path.exists():
        gold_saved_probe = verify_document_enrichment_saved(gold_path, gold_path, spec)
        if gold_saved_probe.error and gold_saved_probe.message.startswith(
            "ERROR: Gold saved document enrichment"
        ):
            return gold_saved_probe
    sub_results = {}
    failures = []
    errors = []

    # Verify footnotes
    if "FootnoteInsert" in commands:
        fn_result = families_notes_verify.verify_footnotes(
            pred_path,
            gold_path,
            spec,
            doc_in_path=doc_in_path,
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        )
        sub_results["footnotes"] = summarize_sub_result(fn_result)
        if fn_result.error:
            errors.append(fn_result.error)
        elif not fn_result.passed:
            failures.append(fn_result.message)

    # Verify endnotes
    if "EndnoteInsertWord" in commands:
        en_result = families_notes_verify.verify_endnotes(
            pred_path,
            gold_path,
            spec,
            doc_in_path=doc_in_path,
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        )
        sub_results["endnotes"] = summarize_sub_result(en_result)
        if en_result.error:
            errors.append(en_result.error)
        elif not en_result.passed:
            failures.append(en_result.message)

    # Verify header/footer content
    has_header_footer = any(
        c in commands
        for c in [
            "HeaderInsertGallery",
            "FooterInsertGallery",
            "PageNumbersInHeaderInsertGallery",
            "PageNambersInFooterInsertGallery",
        ]
    )
    if has_header_footer:
        hf_result = families_headers_verify.verify_header_footer_composition(
            pred_path, gold_path, read_header_footer_content=read_header_footer_content
        )
        sub_results["header_footer"] = summarize_sub_result(hf_result)
        if hf_result.error:
            errors.append(hf_result.error)
        elif not hf_result.passed:
            failures.append(hf_result.message)

    # Verify citation
    if "CitationInsert" in commands:
        cit_result = families_references_verify.verify_citation(
            pred_path, gold_path, spec, read_dynamic_reference_state=read_dynamic_reference_state
        )
        sub_results["citation"] = summarize_sub_result(cit_result)
        if cit_result.error:
            errors.append(cit_result.error)
        elif not cit_result.passed:
            failures.append(cit_result.message)

    # Verify bibliography using the subcommand identity rather than the
    # outer composition id, so composition cases inherit atomic semantics.
    bibliography_command = None
    if "BibliographyInsert" in commands:
        bibliography_command = "BibliographyInsert"
    elif "BibliographyGallery" in commands:
        bibliography_command = "BibliographyGallery"

    if bibliography_command:
        bib_spec = dict(spec) if spec else {}
        bib_spec["idMso"] = bibliography_command
        bib_result = families_references_bibliography.verify_bibliography(
            pred_path, gold_path, bib_spec, read_bibliography_info=read_bibliography_info
        )
        sub_results["bibliography"] = summarize_sub_result(bib_result)
        if bib_result.error:
            errors.append(bib_result.error)
        elif not bib_result.passed:
            failures.append(bib_result.message)

    details = {"sub_results": sub_results}
    atomic_failures, atomic_errors = add_composition_atomic_details(
        details, pred_path, gold_path, spec
    )
    failures.extend(atomic_failures)
    errors.extend(atomic_errors)

    preservation_result = (
        families_document_compositions_preservation.verify_document_enrichment_preservation(
            pred_path, gold_path, spec, doc_in_path=doc_in_path
        )
    )
    details["preservation"] = summarize_sub_result(preservation_result)
    if preservation_result.error:
        errors.append(preservation_result.error)
    elif not preservation_result.passed:
        failures.append(preservation_result.message)

    if errors:
        error = "; ".join(errors)
        return VerificationResult(
            passed=False,
            message=(
                "ERROR: Document enrichment composition could not be fully "
                f"evaluated ({len(errors)} error(s))"
            ),
            details={**details, "errors": errors},
            error=error,
        )

    if failures:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Document enrichment composition ({len(failures)} failed check(s))",
            details=details,
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Document enrichment composition ({len(commands)} operations)",
        details=details,
    )
