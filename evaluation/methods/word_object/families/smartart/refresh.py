"""families / smartart / refresh with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Any, Dict, List, Optional


def verify_refreshed_smartart_projection(
    pred_path: Path,
    gold_path: Path,
    gold_semantic: Dict[str, Any],
    pred_semantic: Dict[str, Any],
    raw_cache_trigger: Dict[str, Any],
    idMso: str,
    *,
    extract_refreshed_smartart_projection,
    get_cleanup_word,
) -> VerificationResult:
    """Resolve a structural raw-cache mismatch through temporary SaveAs.

    Hierarchy, sibling order, node type, and semantic text have already
    matched through normal-open Word COM when this helper runs.  The
    refreshed projection therefore owns only the effective visible node
    sequence/count and page count; GUIDs, raw element order, responsive
    coordinates, and whole-package equality are deliberately excluded.
    """

    # The semantic extraction above owns a long-lived Word instance. Shut
    # it down before the dedicated refresh observations so COM-backed
    # evaluation remains strictly sequential, including within one job.
    cleanup_word = get_cleanup_word()
    if callable(cleanup_word) and cleanup_word() is False:
        message = "Semantic SmartArt Word process did not exit before refresh"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                "raw_cache_trigger": raw_cache_trigger,
                "idMso": idMso,
                "error_owner": "evaluator_runtime",
                "verification_source": "word_temp_saveas_projection",
            },
            error=message,
        )
    gold_projection = extract_refreshed_smartart_projection(gold_path)
    pred_projection = extract_refreshed_smartart_projection(pred_path)
    evidence = {
        "raw_cache_trigger": raw_cache_trigger,
        "gold_projection": gold_projection,
        "pred_projection": pred_projection,
    }

    if not gold_projection.get("found"):
        message = "Gold SmartArt refresh projection could not be established"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                **evidence,
                "idMso": idMso,
                "error_owner": "gold_fixture_or_evaluator",
                "verification_source": "word_temp_saveas_projection",
            },
            error=message,
        )
    if not pred_projection.get("found"):
        message = "Prediction SmartArt refresh projection could not be " "established"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                **evidence,
                "idMso": idMso,
                "error_owner": "evaluator_refresh_observation",
                "verification_source": "word_temp_saveas_projection",
            },
            error=message,
        )

    gold_visible = list(gold_projection.get("visible_node_order") or [])
    pred_visible = list(pred_projection.get("visible_node_order") or [])
    gold_expected = [
        str(text).strip() for text in gold_semantic.get("all_node_texts") or [] if str(text).strip()
    ]
    pred_expected = [
        str(text).strip() for text in pred_semantic.get("all_node_texts") or [] if str(text).strip()
    ]

    if not gold_expected or sorted(gold_visible) != sorted(gold_expected):
        message = "Gold SmartArt refreshed visible nodes do not cover its " "semantic nodes"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                **evidence,
                "idMso": idMso,
                "error_owner": "gold_fixture_or_evaluator",
                "verification_source": "word_temp_saveas_projection",
            },
            error=message,
        )

    differences = []
    if sorted(pred_visible) != sorted(pred_expected):
        differences.append(
            {
                "property": "refreshed_visible_node_membership",
                "gold": sorted(gold_expected),
                "pred": sorted(pred_visible),
                "pred_semantic": sorted(pred_expected),
            }
        )
    if gold_visible != pred_visible:
        differences.append(
            {
                "property": "refreshed_visible_node_order",
                "gold": gold_visible,
                "pred": pred_visible,
            }
        )
    if gold_projection.get("visible_node_count") != pred_projection.get("visible_node_count"):
        differences.append(
            {
                "property": "refreshed_visible_node_count",
                "gold": gold_projection.get("visible_node_count"),
                "pred": pred_projection.get("visible_node_count"),
            }
        )

    gold_pages = gold_projection.get("page_count")
    pred_pages = pred_projection.get("page_count")
    if gold_pages is None or pred_pages is None:
        message = "SmartArt refresh projection lacks Word page-count evidence"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                **evidence,
                "idMso": idMso,
                "error_owner": "evaluator_refresh_observation",
                "verification_source": "word_temp_saveas_projection",
            },
            error=message,
        )
    if gold_pages != pred_pages:
        differences.append(
            {
                "property": "refreshed_page_count",
                "gold": gold_pages,
                "pred": pred_pages,
            }
        )

    if differences:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Refreshed SmartArt visible projection differs "
                f"({len(differences)} difference(s))"
            ),
            details={
                **evidence,
                "differences": differences,
                "idMso": idMso,
                "match_mode": "refresh_aware_word_visible_projection",
                "verification_source": "word_temp_saveas_projection",
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Refreshed SmartArt visible projection matches",
        details={
            **evidence,
            "differences": [],
            "idMso": idMso,
            "match_mode": "refresh_aware_word_visible_projection",
            "verification_source": "word_temp_saveas_projection",
        },
    )


def verify_refreshed_smartart_gallery_projection(
    pred_path: Path,
    gold_path: Path,
    gold_semantic: Dict[str, Any],
    pred_semantic: Dict[str, Any],
    raw_cache_trigger: Dict[str, Any],
    idMso: str,
    *,
    extract_refreshed_smartart_projection,
    get_cleanup_word,
) -> VerificationResult:
    """Adjudicate a stale style/color cache through isolated Word SaveAs.

    The caller has already established matching named gallery identity via
    normal-open Word COM.  This helper does not replace that primary
    observation.  It only resolves a conflicting raw ``drawing*.xml``
    cache by requiring both the refreshed requested gallery type ID and the
    refreshed visible node order/count/styles and Word page count to match
    gold.  Thus stale style markup cannot rescue collateral SmartArt text
    or pagination changes.  Any missing refresh, source mutation, or
    process-cleanup evidence remains indeterminate.
    """

    cleanup_word = get_cleanup_word()
    if callable(cleanup_word) and cleanup_word() is False:
        message = "Semantic SmartArt Word process did not exit before gallery " "refresh"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                "raw_cache_trigger": raw_cache_trigger,
                "word_com_primary": {
                    "gold": gold_semantic,
                    "pred": pred_semantic,
                },
                "idMso": idMso,
                "error_owner": "evaluator_runtime",
                "verification_source": "word_temp_saveas_gallery_projection",
            },
            error=message,
        )

    try:
        gold_projection = extract_refreshed_smartart_projection(gold_path)
    except Exception as exc:
        message = f"Gold SmartArt gallery refresh failed: {exc}"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                "raw_cache_trigger": raw_cache_trigger,
                "word_com_primary": {
                    "gold": gold_semantic,
                    "pred": pred_semantic,
                },
                "idMso": idMso,
                "error_owner": "gold_fixture_or_evaluator",
                "verification_source": "word_temp_saveas_gallery_projection",
            },
            error=message,
        )
    try:
        pred_projection = extract_refreshed_smartart_projection(pred_path)
    except Exception as exc:
        message = f"Prediction SmartArt gallery refresh failed: {exc}"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                "raw_cache_trigger": raw_cache_trigger,
                "gold_projection": gold_projection,
                "word_com_primary": {
                    "gold": gold_semantic,
                    "pred": pred_semantic,
                },
                "idMso": idMso,
                "error_owner": "evaluator_refresh_observation",
                "verification_source": "word_temp_saveas_gallery_projection",
            },
            error=message,
        )

    evidence = {
        "raw_cache_trigger": raw_cache_trigger,
        "gold_projection": gold_projection,
        "pred_projection": pred_projection,
        "word_com_primary": {
            "gold": gold_semantic,
            "pred": pred_semantic,
        },
    }

    def projection_error(
        projection: Dict[str, Any],
    ) -> Optional[str]:
        if not projection.get("found"):
            return str(
                projection.get("error") or "refreshed SmartArt drawing projection is unavailable"
            )
        if projection.get("source_unchanged") is not True:
            return "canonical source hash was not proven unchanged"
        if projection.get("refresh_process_exited") is not True:
            return "dedicated refresh Word process did not exit cleanly"
        if not projection.get("gallery_identities_found"):
            return "refreshed SmartArt gallery identity is unavailable"
        if not projection.get("visible_node_order"):
            return "refreshed SmartArt visible node order is unavailable"
        if projection.get("visible_node_count") is None:
            return "refreshed SmartArt visible node count is unavailable"
        if not projection.get("visible_node_styles"):
            return "refreshed SmartArt visible node styles are unavailable"
        if projection.get("page_count") is None:
            return "refreshed Word page count is unavailable"
        return None

    gold_error = projection_error(gold_projection)
    if gold_error:
        message = f"Gold SmartArt gallery refresh is indeterminate: {gold_error}"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                **evidence,
                "idMso": idMso,
                "error_owner": "gold_fixture_or_evaluator",
                "verification_source": "word_temp_saveas_gallery_projection",
            },
            error=message,
        )
    pred_error = projection_error(pred_projection)
    if pred_error:
        message = "Prediction SmartArt gallery refresh is indeterminate: " f"{pred_error}"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                **evidence,
                "idMso": idMso,
                "error_owner": "evaluator_refresh_observation",
                "verification_source": "word_temp_saveas_gallery_projection",
            },
            error=message,
        )

    identity_property = {
        "SmartArtStylesGallery": "quick_style_type_id",
        "SmartArtChangeColorsGallery": "color_type_id",
    }[idMso]

    def identity_values(projection: Dict[str, Any]) -> List[str]:
        return sorted(
            str(row.get(identity_property) or "")
            for row in projection.get("gallery_identities") or []
            if str(row.get(identity_property) or "")
        )

    gold_identities = identity_values(gold_projection)
    pred_identities = identity_values(pred_projection)
    if not gold_identities:
        message = "Gold SmartArt gallery refresh lacks the requested semantic " "identity"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {message}",
            details={
                **evidence,
                "idMso": idMso,
                "requested_identity_property": identity_property,
                "error_owner": "gold_fixture_or_evaluator",
                "verification_source": "word_temp_saveas_gallery_projection",
            },
            error=message,
        )

    differences = []
    if gold_identities != pred_identities:
        differences.append(
            {
                "property": f"refreshed_{identity_property}",
                "gold": gold_identities,
                "pred": pred_identities,
            }
        )
    gold_styles = list(gold_projection.get("visible_node_styles") or [])
    pred_styles = list(pred_projection.get("visible_node_styles") or [])
    gold_order = list(gold_projection.get("visible_node_order") or [])
    pred_order = list(pred_projection.get("visible_node_order") or [])
    if gold_order != pred_order:
        differences.append(
            {
                "property": "refreshed_smartart_visible_node_order",
                "gold": gold_order,
                "pred": pred_order,
            }
        )
    if gold_projection.get("visible_node_count") != pred_projection.get("visible_node_count"):
        differences.append(
            {
                "property": "refreshed_smartart_visible_node_count",
                "gold": gold_projection.get("visible_node_count"),
                "pred": pred_projection.get("visible_node_count"),
            }
        )
    if gold_styles != pred_styles:
        differences.append(
            {
                "property": "refreshed_smartart_drawing_node_styles",
                "gold": gold_styles,
                "pred": pred_styles,
            }
        )
    if gold_projection.get("page_count") != pred_projection.get("page_count"):
        differences.append(
            {
                "property": "refreshed_page_count",
                "gold": gold_projection.get("page_count"),
                "pred": pred_projection.get("page_count"),
            }
        )

    if differences:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Refreshed SmartArt gallery projection differs "
                f"({len(differences)} difference(s))"
            ),
            details={
                **evidence,
                "differences": differences,
                "idMso": idMso,
                "requested_identity_property": identity_property,
                "match_mode": "refresh_aware_word_gallery_projection",
                "failure_owner": "prediction_saved_state",
                "verification_source": "word_temp_saveas_gallery_projection",
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Refreshed SmartArt gallery projection matches",
        details={
            **evidence,
            "differences": [],
            "idMso": idMso,
            "requested_identity_property": identity_property,
            "match_mode": "refresh_aware_word_gallery_projection",
            "verification_source": "word_com+word_temp_saveas_gallery_guardrail",
        },
    )
