"""composition_content with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.core import DocumentOpenError
from pathlib import Path
from typing import Any, Dict, List, Optional
import re


def summarize_sub_result(result: VerificationResult) -> Dict[str, Any]:
    return {
        "passed": result.passed,
        "message": result.message,
        "details": result.details,
        "error": result.error,
    }


def add_composition_atomic_details(
    details: Dict[str, Any],
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict],
    *,
    read_comment_info,
    read_document_text,
    read_header_footer_content,
    read_revision_info,
    verify_bibliography,
    verify_citation,
    verify_endnotes,
    verify_footnotes,
    verify_header_footer,
) -> tuple[List[str], List[str]]:
    atomic_results = build_composition_atomic_results(
        pred_path,
        gold_path,
        spec,
        read_comment_info=read_comment_info,
        read_document_text=read_document_text,
        read_header_footer_content=read_header_footer_content,
        read_revision_info=read_revision_info,
        verify_bibliography=verify_bibliography,
        verify_citation=verify_citation,
        verify_endnotes=verify_endnotes,
        verify_footnotes=verify_footnotes,
        verify_header_footer=verify_header_footer,
    )
    if not atomic_results:
        return [], []

    target_results = []
    for intent_result in atomic_results:
        for target_result in intent_result.get("target_results", []):
            flat_target = dict(target_result)
            flat_target["intent_id"] = intent_result.get("intent_id")
            flat_target["operation"] = intent_result.get("operation")
            flat_target["atomic_case"] = intent_result.get("atomic_case")
            target_results.append(flat_target)

    details["atomic_results"] = atomic_results
    details["target_results"] = target_results
    details["atomic_summary"] = {
        "passed_intents": sum(1 for result in atomic_results if result.get("passed")),
        "total_intents": len(atomic_results),
        "passed_targets": sum(1 for result in target_results if result.get("passed")),
        "total_targets": len(target_results),
    }
    failures = []
    errors = []
    for target_result in target_results:
        label = (
            f"{target_result.get('intent_id') or 'unknown_intent'}/"
            f"{target_result.get('operation') or 'unknown_operation'}"
        )
        if target_result.get("error"):
            errors.append(f"Atomic composition target {label}: " f"{target_result['error']}")
        elif not target_result.get("passed"):
            failures.append(
                f"Atomic composition target {label} failed: "
                f"{target_result.get('message') or 'semantic value differs'}"
            )
    return failures, errors


def build_composition_atomic_results(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict],
    *,
    read_comment_info,
    read_document_text,
    read_header_footer_content,
    read_revision_info,
    verify_bibliography,
    verify_citation,
    verify_endnotes,
    verify_footnotes,
    verify_header_footer,
) -> List[Dict[str, Any]]:
    intents = (spec or {}).get("composition", {}).get("atomic_intents", [])
    if not intents:
        return []

    cache: Dict[str, Any] = {}
    atomic_results = []
    for intent in intents:
        target_results = [
            evaluate_composition_target(
                target,
                pred_path,
                gold_path,
                spec,
                cache,
                read_comment_info=read_comment_info,
                read_document_text=read_document_text,
                read_header_footer_content=read_header_footer_content,
                read_revision_info=read_revision_info,
                verify_bibliography=verify_bibliography,
                verify_citation=verify_citation,
                verify_endnotes=verify_endnotes,
                verify_footnotes=verify_footnotes,
                verify_header_footer=verify_header_footer,
            )
            for target in intent.get("verification_targets", []) or []
        ]
        atomic_results.append(
            {
                "intent_id": intent.get("id"),
                "operation": intent.get("operation"),
                "atomic_case": intent.get("atomic_case"),
                "passed": bool(target_results)
                and all(result.get("passed") for result in target_results),
                "target_results": target_results,
            }
        )
    return atomic_results


def evaluate_composition_target(
    target: Dict[str, Any],
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict],
    cache: Dict[str, Any],
    *,
    read_comment_info,
    read_document_text,
    read_header_footer_content,
    read_revision_info,
    verify_bibliography,
    verify_citation,
    verify_endnotes,
    verify_footnotes,
    verify_header_footer,
) -> Dict[str, Any]:
    extractor = target.get("extractor")
    prop = target.get("property")
    result: Dict[str, Any] = {
        "target": target,
        "passed": False,
    }

    if not extractor or not prop:
        result["error"] = "Missing extractor or property"
        return result

    try:
        if extractor == "content_operation_verifier":
            sub_result = verify_content_operation_target(
                prop,
                pred_path,
                gold_path,
                spec,
                target,
                verify_bibliography=verify_bibliography,
                verify_citation=verify_citation,
                verify_endnotes=verify_endnotes,
                verify_footnotes=verify_footnotes,
                verify_header_footer=verify_header_footer,
            )
            result.update(
                {
                    "passed": sub_result.passed,
                    "message": sub_result.message,
                    "details": sub_result.details,
                }
            )
            if sub_result.error:
                result["error"] = sub_result.error
            return result

        gold_data = extract_composition_target_data(
            extractor,
            gold_path,
            cache,
            "gold",
            read_comment_info=read_comment_info,
            read_document_text=read_document_text,
            read_header_footer_content=read_header_footer_content,
            read_revision_info=read_revision_info,
        )
        pred_data = extract_composition_target_data(
            extractor,
            pred_path,
            cache,
            "pred",
            read_comment_info=read_comment_info,
            read_document_text=read_document_text,
            read_header_footer_content=read_header_footer_content,
            read_revision_info=read_revision_info,
        )
        gold_value = composition_target_value(gold_data, prop, target)
        pred_value = composition_target_value(pred_data, prop, target)
        expected = target.get("expected")
        passed = (
            pred_value == expected
            if "expected" in target
            else composition_values_match(gold_value, pred_value)
        )
        result.update(
            {
                "passed": passed,
                "gold": gold_value,
                "pred": pred_value,
            }
        )
        if "expected" in target:
            result["expected"] = expected
        if not passed:
            result["differences"] = [
                {
                    "property": prop,
                    "gold": gold_value,
                    "pred": pred_value,
                    **({"expected": expected} if "expected" in target else {}),
                }
            ]
    except DocumentOpenError:
        raise
    except Exception as exc:
        result.update(
            {
                "passed": False,
                "error": str(exc),
            }
        )
    return result


def verify_content_operation_target(
    prop: str,
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict],
    target: Dict[str, Any],
    *,
    verify_bibliography,
    verify_citation,
    verify_endnotes,
    verify_footnotes,
    verify_header_footer,
) -> VerificationResult:
    scoped = dict(spec or {})
    if target.get("idMso"):
        scoped["idMso"] = target["idMso"]
    if target.get("source_tag"):
        scoped["source_tag"] = target["source_tag"]

    if prop == "footnote_verification":
        return verify_footnotes(pred_path, gold_path, scoped)
    if prop == "endnote_verification":
        return verify_endnotes(pred_path, gold_path, scoped)
    if prop == "citation_verification":
        return verify_citation(pred_path, gold_path, scoped)
    if prop == "bibliography_verification":
        return verify_bibliography(pred_path, gold_path, scoped)
    if prop == "header_footer_style_match":
        return verify_header_footer(
            pred_path,
            gold_path,
            target.get("part_type", "footer"),
            "style_match",
            scoped,
            verify_body_preservation=False,
        )
    if prop == "header_footer_page_field_match":
        return verify_header_footer(
            pred_path,
            gold_path,
            target.get("part_type", "footer"),
            "page_field_match",
            scoped,
            verify_body_preservation=False,
        )
    return VerificationResult(
        passed=False,
        message=f"Unsupported content-operation verifier target: {prop}",
        details={},
    )


def extract_composition_target_data(
    extractor: str,
    path: Path,
    cache: Dict[str, Any],
    side: str,
    *,
    read_comment_info,
    read_document_text,
    read_header_footer_content,
    read_revision_info,
) -> Any:
    key = f"{side}:{extractor}"
    if key in cache:
        return cache[key]

    if extractor == "header_footer_content":
        data = read_header_footer_content(str(path))
    elif extractor == "revision_info":
        data = read_revision_info(str(path))
    elif extractor == "document_text":
        data = read_document_text(str(path)).strip()
    elif extractor == "comment_info":
        data = read_comment_info(str(path))
    else:
        raise ValueError(f"Unsupported content-operation target extractor: {extractor}")

    cache[key] = data
    return data


def composition_target_value(data: Any, prop: str, target: Optional[Dict[str, Any]] = None) -> Any:
    target = target or {}
    if isinstance(data, str):
        return data if prop == "text" else None
    if isinstance(data, dict):
        if "error" in data:
            raise ValueError(data["error"])
        if "part" in target and ("headers" in data or "footers" in data):
            return header_footer_target_value(data, prop, target["part"])
        return data.get(prop)
    return None


def header_footer_target_value(data: Dict[str, Any], prop: str, part: str) -> Any:
    parts = data.get("headers" if part == "header" else "footers", []) or []

    def _normalize_text(text: str) -> str:
        return re.sub(r"\s+", " ", (text or "").replace("\r", " ")).strip()

    if prop == "text":
        return sorted(_normalize_text(item.get("text", "")) for item in parts)
    if prop in {
        "field_count",
        "page_field_count",
        "num_pages_field_count",
        "inline_shape_count",
        "shape_count",
    }:
        return sorted(item.get(prop, 0) or 0 for item in parts)
    if prop == "field_keywords":
        keywords = []
        for item in parts:
            for code in item.get("field_codes", []) or []:
                normalized = " ".join((code or "").upper().split())
                if "NUMPAGES" in normalized:
                    keywords.append("NUMPAGES")
                elif "PAGE" in normalized:
                    keywords.append("PAGE")
                elif normalized:
                    keywords.append(normalized.split(" ")[0])
        return sorted(keywords)
    raise ValueError(f"Unsupported header/footer target property: {prop}")


def composition_values_match(gold_value: Any, pred_value: Any) -> bool:
    return gold_value == pred_value
