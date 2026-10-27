"""families / equations / verify with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.core import DocumentOpenError
from evaluation.methods.word_object.families.equations import compare as families_equations_compare
from evaluation.methods.word_object.families.equations import (
    extract_saved as families_equations_extract_saved,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import unicodedata


def _text_only_preset(spec, group):
    """Public oracles retain variation names, without generator parameters."""
    command, variants = {
        "basic_operator": ("EquationOperatorGallery", {"colon_equals", "def_equals"}),
        "differential": ("EquationIntegralGallery", {"dx", "dy", "d_theta"}),
    }[group]
    spec = spec or {}
    parameter = spec.get("parameter") or {}
    return spec.get("idMso") == command and (
        spec.get("variation") in variants
        or (parameter.get("preset_group") == group and parameter.get("omath_function") is None)
    )


def verify_equation_com_special(
    pred_path: Path, gold_path: Path, *, read_equation_structure,
    allow_compact_cases: bool = False,
    allow_logarithm_name: bool = False,
) -> VerificationResult:
    gold = read_equation_structure(str(gold_path))
    pred = read_equation_structure(str(pred_path))
    signature = families_equations_compare.equation_com_semantic_signature
    if allow_compact_cases:
        signature = families_equations_compare.compact_cases_com_signature
    elif allow_logarithm_name:
        signature = families_equations_compare.logarithm_name_com_signature
    gold_signature = signature(gold)
    pred_signature = signature(pred)
    return VerificationResult(
        passed=gold_signature == pred_signature,
        message=(
            "PASS: Word Object Model equation semantics match"
            if gold_signature == pred_signature
            else "FAIL: Word Object Model equation semantics differ"
        ),
        details={
            "gold": gold_signature,
            "pred": pred_signature,
            **({"match_mode": "typed_compact_cases_rows"} if allow_compact_cases else {}),
            **({"match_mode": "typed_logarithm_operator_name"} if allow_logarithm_name else {}),
        },
    )


def verify_equation_saved_special(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, saved_evidence_error
) -> VerificationResult:
    interaction_class = (spec or {}).get("composition", {}).get("interaction_class", "")
    id_mso = (spec or {}).get("idMso", "")
    parameter = (spec or {}).get("parameter") or {}
    variation = (spec or {}).get("variation")
    if id_mso == "EquationLimitGallery" and (
        parameter.get("label") == "logarithm_with_empty_base"
        or variation == "logarithm_with_empty_base"
    ):
        expected = [
            {
                "function_name": "log",
                "subscript": "2",
                "argument": "x",
            }
        ]
        try:
            gold_signature = (
                families_equations_extract_saved.extract_saved_logarithm_with_base_signature(
                    gold_path
                )
            )
        except Exception as exc:
            return saved_evidence_error(
                f"Gold saved logarithm extraction failed: {exc}",
                evidence_owner="gold",
                exception=exc,
            )
        if gold_signature != expected:
            return saved_evidence_error(
                "Gold lacks the canonical native m:sSub(log, 2) " "function with x argument",
                evidence_owner="gold",
            )
        try:
            pred_signature = (
                families_equations_extract_saved.extract_saved_logarithm_with_base_signature(
                    pred_path
                )
            )
        except Exception as exc:
            return saved_evidence_error(
                f"Prediction saved logarithm extraction failed: {exc}",
                evidence_owner="prediction",
                exception=exc,
            )
        return VerificationResult(
            passed=pred_signature == gold_signature,
            message=(
                "PASS: Saved logarithm-with-base OMML structure matches"
                if pred_signature == gold_signature
                else "FAIL: Saved logarithm-with-base OMML structure differs"
            ),
            details={
                "match_mode": "saved_omml_logarithm_with_base",
                "gold": gold_signature,
                "pred": pred_signature,
            },
        )
    if id_mso == "EquationDelimiterGallery" and (
        parameter.get("label") == "three_conditions" or variation == "three_conditions"
    ):
        expected = [
            {
                "prefix_text": "y=",
                "container_kind": "eqArr",
                "delimiter_begin": "{",
                "delimiter_end": "",
                "rows": ["a", "b", "c"],
                "base_justification": "center",
                "row_spacing": 0,
                "row_spacing_rule": 0,
                "maximum_distance": False,
                "object_distance": False,
            }
        ]
        try:
            gold_signature = (
                families_equations_extract_saved.extract_saved_three_conditions_signature(gold_path)
            )
        except Exception as exc:
            return saved_evidence_error(
                f"Gold saved three-condition extraction failed: {exc}",
                evidence_owner="gold",
                exception=exc,
            )
        if gold_signature != expected:
            return saved_evidence_error(
                "Gold lacks the canonical compact three-condition " "m:eqArr layout",
                evidence_owner="gold",
            )
        try:
            pred_signature = (
                families_equations_extract_saved.extract_saved_three_conditions_signature(pred_path)
            )
        except Exception as exc:
            return saved_evidence_error(
                f"Prediction saved three-condition extraction failed: {exc}",
                evidence_owner="prediction",
                exception=exc,
            )

        def semantic_layout(
            signatures: List[Dict[str, Any]],
        ) -> List[Dict[str, Any]]:
            """Normalize render-equivalent one-column matrix cases.

            Word preserves some visually identical three-row cases as a
            one-column ``m:m`` rather than the gallery's ``m:eqArr``.
            Container identity alone is not render-critical: the row
            topology, delimiter, alignment, and spacing still are.
            """
            normalized = []
            for signature in signatures:
                kind = signature.get("container_kind")
                rows = signature.get("rows")
                if kind == "m":
                    if not (
                        isinstance(rows, list)
                        and all(isinstance(row, list) and len(row) == 1 for row in rows)
                    ):
                        normalized.append(
                            {
                                **signature,
                                "container_equivalence": "non_equivalent_matrix",
                            }
                        )
                        continue
                    rows = [row[0] for row in rows]
                semantic = {
                    key: value for key, value in signature.items() if key != "container_kind"
                }
                semantic.update(
                    {
                        "rows": rows,
                        "container_equivalence": "compact_cases_rows",
                    }
                )
                normalized.append(semantic)
            return normalized

        gold_semantic = semantic_layout(gold_signature)
        pred_semantic = semantic_layout(pred_signature)
        passed = pred_semantic == gold_semantic
        return VerificationResult(
            passed=passed,
            message=(
                "PASS: Saved three-condition OMML layout is equivalent"
                if passed
                else "FAIL: Saved three-condition OMML layout differs"
            ),
            details={
                "match_mode": "saved_omml_three_conditions_layout",
                "gold": gold_signature,
                "pred": pred_signature,
                "semantic_gold": gold_semantic,
                "semantic_pred": pred_semantic,
            },
        )
    if id_mso == "EquationOperatorGallery":
        try:
            gold_signature = (
                families_equations_extract_saved.extract_saved_equation_operator_signature(
                    gold_path
                )
            )
        except Exception as exc:
            return saved_evidence_error(
                f"Gold saved equation extraction failed: {exc}",
                evidence_owner="gold",
                exception=exc,
            )
        if not gold_signature:
            return saved_evidence_error(
                "Gold saved equation-operator evidence is empty",
                evidence_owner="gold",
            )
        try:
            pred_signature = (
                families_equations_extract_saved.extract_saved_equation_operator_signature(
                    pred_path
                )
            )
        except Exception as exc:
            return saved_evidence_error(
                f"Prediction saved equation extraction failed: {exc}",
                evidence_owner="prediction",
                exception=exc,
            )
        return VerificationResult(
            passed=gold_signature == pred_signature,
            message=(
                "PASS: Saved equation operator signature matches"
                if gold_signature == pred_signature
                else "FAIL: Saved equation operator signature differs"
            ),
            details={
                "match_mode": "saved_omml_operator",
                "gold": gold_signature,
                "pred": pred_signature,
            },
        )
    if interaction_class == "insert_delimiter_then_matrix":
        try:
            gold_signature = families_equations_extract_saved.extract_delimiter_matrix_signature(
                gold_path
            )
        except Exception as exc:
            return saved_evidence_error(
                f"Gold saved delimiter/matrix extraction failed: {exc}",
                evidence_owner="gold",
                exception=exc,
            )
        if (
            not gold_signature.get("equation_count")
            or gold_signature.get("delimiter") is None
            or gold_signature.get("matrix") is None
        ):
            return saved_evidence_error(
                "Gold saved delimiter/matrix evidence is incomplete",
                evidence_owner="gold",
            )
        try:
            pred_signature = families_equations_extract_saved.extract_delimiter_matrix_signature(
                pred_path
            )
        except Exception as exc:
            return saved_evidence_error(
                f"Prediction saved delimiter/matrix extraction failed: {exc}",
                evidence_owner="prediction",
                exception=exc,
            )
        return VerificationResult(
            passed=gold_signature == pred_signature,
            message=(
                "PASS: Delimiter/matrix OMML semantic signature matches"
                if gold_signature == pred_signature
                else "FAIL: Delimiter/matrix OMML semantic signature differs"
            ),
            details={
                "match_mode": "saved_omml_delimiter_matrix",
                "gold": gold_signature,
                "pred": pred_signature,
            },
        )
    return VerificationResult.error_result(
        "No saved equation verifier is registered for this command"
    )


def verify_equation(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_equation_structure,
    saved_evidence_error,
) -> VerificationResult:
    """
    Verify equation structure matches.

    Handles EquationAccentGallery ~ EquationSymbolsInsertGallery.
    Compares function count, function types, and equation text.
    """
    interaction_class = (spec or {}).get("composition", {}).get("interaction_class", "")
    idMso = (spec or {}).get("idMso", "")
    equation_parameter = (spec or {}).get("parameter") or {}
    equation_variation = (spec or {}).get("variation")
    is_text_only_basic_operator = _text_only_preset(spec, "basic_operator")
    is_logarithm_with_base = idMso == "EquationLimitGallery" and (
        equation_parameter.get("label") == "logarithm_with_empty_base"
        or equation_variation == "logarithm_with_empty_base"
    )
    is_three_conditions = idMso == "EquationDelimiterGallery" and (
        equation_parameter.get("label") == "three_conditions"
        or (spec or {}).get("variation") == "three_conditions"
    )
    if (
        (idMso == "EquationOperatorGallery" and not is_text_only_basic_operator)
        or interaction_class == "insert_delimiter_then_matrix"
        or is_logarithm_with_base
        or is_three_conditions
    ):
        saved_result = verify_equation_saved_special(
            pred_path, gold_path, spec, saved_evidence_error=saved_evidence_error
        )
        if (
            (is_logarithm_with_base or is_three_conditions)
            and saved_result.error
            and saved_result.details.get("evidence_owner") == "gold"
        ):
            # A case-specific structural contract cannot be adjudicated
            # when its gold fixture lacks that structure, regardless of
            # what the higher-level COM projection happens to report.
            return tag_verification_source(saved_result, "saved_ooxml_guardrail")
        try:
            com_result = verify_equation_com_special(
                pred_path, gold_path, read_equation_structure=read_equation_structure,
                allow_compact_cases=is_three_conditions and saved_result.passed,
                allow_logarithm_name=is_logarithm_with_base and saved_result.passed,
            )
        except DocumentOpenError:
            raise
        except Exception as com_error:
            if saved_result.error or not saved_result.passed:
                # A saved mismatch is sufficient negative evidence.  An
                # extraction error also remains indeterminate.
                return tag_verification_source(
                    saved_result,
                    "saved_ooxml_fallback",
                    word_com_error=str(com_error),
                )
            error = (
                "Word COM equation semantics are unavailable; matching "
                "saved-state structure does not establish the complete "
                "equation and preservation contract"
            )
            return tag_verification_source(
                VerificationResult(
                    passed=False,
                    message=f"ERROR: {error}",
                    details={
                        "saved_ooxml_evidence": saved_result.to_dict(),
                        "missing_evidence": "word_com_equation_semantics",
                    },
                    error=error,
                ),
                "saved_ooxml_fallback",
                word_com_error=str(com_error),
            )
        if not com_result.passed:
            return tag_verification_source(com_result, "word_com")
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_result)

    gold_eq = read_equation_structure(str(gold_path))
    pred_eq = read_equation_structure(str(pred_path))

    # Legacy top-level fields describe only equation 1. Other equations
    # remain preservation evidence and must not disappear from comparison.
    if not gold_eq.get('equation_count'):
        return VerificationResult.error_result('Gold equation evidence is missing')
    for owner, state in (("Gold", gold_eq), ("Prediction", pred_eq)):
        if 'equations' in state and len(state['equations']) != state.get('equation_count'):
            return VerificationResult.error_result(f'{owner} equation observation is incomplete')
    gold_tail = gold_eq.get('equations', [])[1:]
    pred_tail = pred_eq.get('equations', [])[1:]
    if [families_equations_compare.equation_com_semantic_signature(item) for item in gold_tail] != [
        families_equations_compare.equation_com_semantic_signature(item) for item in pred_tail
    ]:
        return VerificationResult(passed=False, message='FAIL: Non-target equations changed',
                                  details={'gold_equations':gold_tail, 'pred_equations':pred_tail})

    def _normalized_text(equation: Dict[str, Any]) -> str:
        # COM inserts paragraph/run separators around nested OMML sSup
        # functions, whereas Unicode superscripts stay in one run.
        return families_equations_compare.normalized_equation_text(equation)

    equivalence = families_equations_compare.equation_contract_equivalence(spec or {})
    equivalence_details = {}
    if equivalence:
        equivalence_details = {"comparison_equivalence": equivalence,
                               "gold_observation": gold_eq, "pred_observation": pred_eq}
        gold_eq = families_equations_compare.normalize_equation_contract(gold_eq, equivalence)
        pred_eq = families_equations_compare.normalize_equation_contract(pred_eq, equivalence)

    if idMso == "EquationRadicalGallery":
        # Word can encode exponents inside a radical either as Unicode
        # superscript characters or nested OMML sSup functions. The COM
        # function count differs, but folded NFKC text plus radical degree
        # mode captures the same requested mathematical object. The lax
        # fold is safe here only because the degree-mode lists are also
        # compared; see _folded_equation_text.
        gold_radicals = [func for func in gold_eq.get("functions", []) if "rad_hide_degree" in func]
        pred_radicals = [func for func in pred_eq.get("functions", []) if "rad_hide_degree" in func]
        if (
            gold_eq.get("equation_count") == pred_eq.get("equation_count")
            and families_equations_compare.folded_equation_text(gold_eq)
            == families_equations_compare.folded_equation_text(pred_eq)
            and [func.get("rad_hide_degree") for func in gold_radicals]
            == [func.get("rad_hide_degree") for func in pred_radicals]
            and gold_radicals
            and (gold_eq.get('semantic_expression') == pred_eq.get('semantic_expression'))
        ):
            return VerificationResult(
                passed=True,
                message="PASS: Equation radical is semantically equivalent",
                details={
                    "match_mode": "nfkc_text_plus_radical_degree_mode",
                    "equation_text": families_equations_compare.folded_equation_text(gold_eq)[:50],
                },
            )

    parameter = (spec or {}).get("parameter") or {}
    is_angle_delimiter_request = idMso == "EquationDelimiterGallery" and (
        parameter.get("label") == "angle_brackets"
        or (spec or {}).get("variation") == "angle_brackets"
    )
    if _text_only_preset(spec, "basic_operator"):
        gold_text = families_equations_compare.normalized_basic_operator_text(gold_eq)
        pred_text = families_equations_compare.normalized_basic_operator_text(pred_eq)
        if (
            gold_eq.get("equation_count") == pred_eq.get("equation_count")
            and gold_text == pred_text
        ):
            return VerificationResult(
                passed=True,
                message="PASS: Basic equation operator text is semantically equivalent",
                details={
                    "match_mode": "word_com_normalized_basic_operator_text",
                    "equation_text": gold_text[:50],
                },
            )
    if _text_only_preset(spec, "differential"):
        # Word may wrap inserted differential text in an OMML Box when a
        # user types it through the equation UI. The extra Box is an
        # incidental representation choice: this variant requests text,
        # not a particular equation function. Preserve the WOM text and
        # equation-count checks while ignoring that wrapper.
        if gold_eq.get("equation_count") == pred_eq.get("equation_count") and _normalized_text(
            gold_eq
        ) == _normalized_text(pred_eq):
            return VerificationResult(
                passed=True,
                message="PASS: Equation differential text is semantically equivalent",
                details={
                    "match_mode": "word_com_normalized_differential_text",
                    "equation_text": _normalized_text(gold_eq)[:50],
                    "gold_function_types": [
                        function.get("type") for function in gold_eq.get("functions", [])
                    ],
                    "pred_function_types": [
                        function.get("type") for function in pred_eq.get("functions", [])
                    ],
                },
            )

    differences = []

    # Compare equation count
    if gold_eq.get("equation_count") != pred_eq.get("equation_count"):
        differences.append(
            {
                "property": "equation_count",
                "gold": gold_eq.get("equation_count"),
                "pred": pred_eq.get("equation_count"),
            }
        )

    # Compare function count
    if gold_eq.get("function_count") != pred_eq.get("function_count"):
        differences.append(
            {
                "property": "function_count",
                "gold": gold_eq.get("function_count"),
                "pred": pred_eq.get("function_count"),
            }
        )

    # Compare function types
    gold_funcs = gold_eq.get("functions", [])
    pred_funcs = pred_eq.get("functions", [])

    gold_func_types = [f.get("type") for f in gold_funcs]
    pred_func_types = [f.get("type") for f in pred_funcs]

    if gold_func_types != pred_func_types:
        differences.append(
            {
                "property": "function_types",
                "gold": gold_func_types,
                "pred": pred_func_types,
            }
        )

    # Compare function-specific parameters for matching functions
    for gf, pf in zip(gold_funcs, pred_funcs):
        if gf.get("type") != pf.get("type"):
            continue
        # Compare all extra keys beyond index/type
        for key in gf:
            if key in ("index", "type"):
                continue
            if (
                key == "nary_sub_sup_lim"
                and gf.get("nary_hide_sub")
                and gf.get("nary_hide_sup")
                and pf.get("nary_hide_sub")
                and pf.get("nary_hide_sup")
            ):
                # Limit placement has no visible or semantic effect when
                # both the subscript and superscript are hidden.
                continue
            gold_value = gf.get(key)
            pred_value = pf.get(key)
            if is_angle_delimiter_request and key in {"delim_begin_char", "delim_end_char"}:
                gold_value = families_equations_compare.canonical_angle_bracket(gold_value)
                pred_value = families_equations_compare.canonical_angle_bracket(pred_value)
            if gold_value != pred_value:
                differences.append(
                    {
                        "property": f"func_{gf.get('index')}_{key}",
                        "gold": gold_value,
                        "pred": pred_value,
                    }
                )

    # Compare equation text. NFKC folds Word's auto-italicised Mathematical
    # Italic codepoints (e.g. '𝑦' U+1D466) back to ASCII so visually-equivalent
    # equations compare equal across Word's equation editor and other paths.
    is_angle_delimiter = (
        is_angle_delimiter_request
        and families_equations_compare.has_native_angle_delimiter(gold_eq)
        and families_equations_compare.has_native_angle_delimiter(pred_eq)
    )
    if is_angle_delimiter:
        gold_text = families_equations_compare.normalized_angle_delimiter_text(gold_eq)
        pred_text = families_equations_compare.normalized_angle_delimiter_text(pred_eq)
    else:
        gold_text = _normalized_text(gold_eq)
        pred_text = _normalized_text(pred_eq)
    if gold_text != pred_text:
        differences.append(
            {
                "property": "equation_text",
                "gold": gold_text[:100],
                "pred": pred_text[:100],
            }
        )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Equation mismatch ({len(differences)} difference(s))",
            details={
                "differences": differences,
                "gold_eq": gold_eq,
                "pred_eq": pred_eq,
                **equivalence_details,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Equation structure matches",
        details={
            "function_count": gold_eq.get("function_count"),
            "equation_text": gold_text[:50],
            **equivalence_details,
        },
    )


def verify_equation_conversion(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_equation_structure
) -> VerificationResult:
    """Verify equation conversion commands with multi-equation discrimination.

    Uses per-equation function_count to distinguish All vs One:
    - All commands: ALL equations must match expected format
    - One commands: only FIRST equation matches, others stay opposite
    - NormalText: compare first equation text (ASCII vs math italic)
    """
    gold_eq = read_equation_structure(str(gold_path))
    pred_eq = read_equation_structure(str(pred_path))

    idMso = spec.get("idMso", "") if spec else ""
    variation = spec.get("variation", "on") if spec else "on"

    # Basic check: both must have equations
    if gold_eq.get("equation_count", 0) == 0:
        return VerificationResult.error_result('Gold document has no equations')
    if pred_eq.get("equation_count", 0) == 0:
        return VerificationResult(
            passed=False,
            message="FAIL: Pred document has no equations",
            details={"gold": gold_eq, "pred": pred_eq},
        )

    gold_eqs = gold_eq.get("equations", [])
    pred_eqs = pred_eq.get("equations", [])

    for owner, state, entries in (("Gold",gold_eq,gold_eqs),("Prediction",pred_eq,pred_eqs)):
        if len(entries) != state.get('equation_count'):
            return VerificationResult.error_result(f'{owner} equation observation is incomplete')
    if gold_eq.get('equation_count') != pred_eq.get('equation_count'):
        return VerificationResult(passed=False,message='FAIL: Equation count mismatch',
                                  details={'gold':gold_eq,'pred':pred_eq})
    for index, (gold_item, pred_item) in enumerate(zip(gold_eqs, pred_eqs),1):
        is_conversion_target = (
            idMso in {"EquationProfessionalAll", "EquationLinearAll",
                      "EquationProfessionalOne", "EquationLinearOne"}
            and (idMso.endswith("All") or index == 1)
        )
        signature = (families_equations_compare.linear_conversion_com_signature
                     if is_conversion_target
                     else families_equations_compare.equation_com_semantic_signature)
        if signature(gold_item) != signature(pred_item):
            return VerificationResult(passed=False,
                message=f'FAIL: Equation {index} content or native structure differs',
                details={'gold':gold_item,'pred':pred_item})

    # EquationNormalText: compare first equation text. NFKC folds Word's
    # Mathematical Italic codepoints back to ASCII so equations rendered
    # via different code paths still compare equal (see N6 in
    # the composition review findings).
    if idMso == "EquationNormalText":
        gold_text_raw = (
            gold_eqs[0].get("equation_text", "") if gold_eqs else gold_eq.get("equation_text", "")
        )
        pred_text_raw = (
            pred_eqs[0].get("equation_text", "") if pred_eqs else pred_eq.get("equation_text", "")
        )
        gold_text = unicodedata.normalize("NFKC", gold_text_raw or "")
        pred_text = unicodedata.normalize("NFKC", pred_text_raw or "")
        if gold_text != pred_text:
            return VerificationResult(
                passed=False,
                message="FAIL: Equation text mismatch",
                details={"gold": gold_eq, "pred": pred_eq},
            )
        return VerificationResult(
            passed=True,
            message="PASS: EquationNormalText matches",
            details={"gold": gold_eq, "pred": pred_eq},
        )

    # Professional/Linear commands: per-equation gold vs pred comparison
    # Professional: function_count >= 2, Linear: function_count == 1
    if len(gold_eqs) != len(pred_eqs):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Equation count mismatch (gold={len(gold_eqs)}, pred={len(pred_eqs)})",
            details={"gold": gold_eq, "pred": pred_eq},
        )

    min_count = min(len(gold_eqs), len(pred_eqs))
    if min_count == 0:
        return VerificationResult(
            passed=False,
            message="FAIL: No per-equation data to compare",
            details={"gold": gold_eq, "pred": pred_eq},
        )

    mismatches = []
    for i in range(min_count):
        gold_fc = gold_eqs[i].get("function_count", 0)
        pred_fc = pred_eqs[i].get("function_count", 0)
        gold_is_pro = gold_fc >= 2
        pred_is_pro = pred_fc >= 2
        if gold_is_pro != pred_is_pro:
            g_fmt = "Professional" if gold_is_pro else "Linear"
            p_fmt = "Professional" if pred_is_pro else "Linear"
            mismatches.append(f"eq{i+1}: gold={g_fmt}(fc={gold_fc}) pred={p_fmt}(fc={pred_fc})")

    if mismatches:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Equation format mismatch — {'; '.join(mismatches)}",
            details={"gold": gold_eq, "pred": pred_eq},
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: All {min_count} equations match gold format",
        details={"gold": gold_eq, "pred": pred_eq},
    )
