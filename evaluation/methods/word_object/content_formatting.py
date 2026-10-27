"""
Content formatting verification for WordEval.

Handles verification for content_formatting behavior across all scopes:
- core_actionable (Bold, Italic, Strikethrough, etc.)
- core_parameterized (FontSize, FontColor, etc.)
- contextual_actionable, contextual_parameterized, etc.

Verification methods are spec-driven:
- property_based: Compare specific formatting property on target text
- text_preserved: Check document text content is unchanged
- surrounding_text_unchanged: Verify non-target text wasn't modified
"""

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional

from ..base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from .composition import composition_command_ids, composition_operation_ids
from .core import DocumentOpenError, WordObjectVerification


from .families.formatting.text import (
    BULLET_CHAR_EQUIVALENTS,
    _normalize_bullet_char,
    _UNICODE_BULLET_CHARS,
    _LEGACY_FONTS,
    _FONT_DEPENDENT_BULLET_REQUIREMENTS,
    _LEGACY_FONT_KEYS,
    _bullet_chars_equivalent,
    _adjacent_text_run_groups,
    _xml_element_signature,
    _matching_adjacent_run_signatures,
)

from .families.formatting import compare as families_formatting_compare
from .families.formatting import composition as families_formatting_composition
from .families.formatting import extract_saved as families_formatting_extract_saved
from .families.formatting import paragraphs as families_formatting_paragraphs
from .families.formatting import preservation as families_formatting_preservation
from .families.formatting import verify as families_formatting_verify
from .families.formatting import verify_saved as families_formatting_verify_saved
from .families.formatting import visibility as families_formatting_visibility
from .policies import content_formatting as policies_content_formatting


class ContentFormattingVerifier:
    """
    Verifier for content_formatting behavior.

    Uses WordObjectVerification core to extract text formatting properties
    and compare them between predicted and gold documents.
    """

    def __init__(self, core: WordObjectVerification):
        self._core = core

    def verify_saved_fallback(
        self,
        pred_path: Path,
        gold_path: Path,
        spec: Optional[Dict] = None,
    ) -> Optional[VerificationResult]:
        """Return a registered saved-state verifier after a COM timeout."""
        spec = spec or {}
        idMso = spec.get("idMso", "")
        saved_result = None
        if idMso == "LigatureGalleryWord":
            saved_result = self._verify_ligatures(pred_path, gold_path, spec)
        elif idMso == "MultilevelListGallery":
            saved_result = self._verify_saved_multilevel_list(pred_path, gold_path)
        elif idMso == "QuickStylesGallery":
            saved_result = self._verify_quick_styles_saved(pred_path, gold_path, spec)
        elif idMso in {
            "TextFillGradientGallery",
            "TextOutlineDashesGallery",
            "TextHighlightColorPicker",
        }:
            saved_result = self._verify_saved_target_text_effect(pred_path, gold_path, spec)
        if saved_result is None:
            return None
        return self._saved_result_after_com_failure(
            saved_result,
            (
                "Word COM timed out; saved OOXML target evidence matches, "
                "but visible preservation could not be established"
            ),
        )

    _saved_result_after_com_failure = staticmethod(
        policies_content_formatting.saved_result_after_com_failure
    )

    def _com_primary_with_saved_guardrail(self, com_verify, saved_verify):
        return policies_content_formatting.com_primary_with_saved_guardrail(
            com_verify,
            saved_verify,
        )

    def verify(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path=None
    ) -> VerificationResult:
        """
        Route a content_formatting case to the right verifier.

        All routing decisions live here so any reader (or static analysis
        tool) can see the full dispatch in one place. Three layers, in order:

        1. Methods with hard-coded routes (ignore select_text and idMso):
           - `paragraph_spacing_verification` → `_verify_all_paragraphs`
           - `dropcap_verification`           → `_verify_dropcap`

        2. property_based-family methods with idMso-specific handlers
           (these override select_text — the dedicated handler decides scope):
           - `LigatureGalleryWord`  → `_verify_ligatures`
           - `QuickStylesGallery`   → `_verify_quick_styles`

        3. Everything else honors `precondition.select_text`:
           - `'all'`                  → `_verify_all_paragraphs`
           - `'paragraph_N'` (N int)  → `_verify_specific_paragraph`
           - otherwise dispatch by method:
             - `property_based` / `xml_property_check` / `openxml_comparison`
               → `_verify_property_based`
             - any other method (incl. `numbering_property_comparison` and
               unknown labels) → `_verify_formatting_match`

        With no spec, falls back to `_verify_formatting_match(pred, gold)`
        for a baseline document comparison.
        """
        if not spec:
            return self._verify_formatting_match(pred_path, gold_path)

        primary = spec.get("verification", {}).get("primary", {})
        method = primary.get("method", "property_based")
        select_text = spec.get("precondition", {}).get("select_text", "Target")
        idMso = spec.get("idMso", "")

        composition_commands = set(composition_command_ids(spec))
        if (
            spec.get("composition")
            and {"TextOutlineWeightGallery", "TextReflectionGallery"} <= composition_commands
        ):
            return self._com_primary_with_saved_guardrail(
                lambda: self._verify_property_based(pred_path, gold_path, spec),
                lambda: self._verify_saved_font_composition(pred_path, gold_path, spec),
            )
        if idMso in {
            "TextFillGradientGallery",
            "TextOutlineDashesGallery",
            "TextHighlightColorPicker",
        }:
            return self._com_primary_with_saved_guardrail(
                lambda: self._verify_property_based(pred_path, gold_path, spec),
                lambda: self._verify_saved_target_text_effect(pred_path, gold_path, spec),
            )

        # Layer 1: hard-coded method routes.
        if method == "paragraph_spacing_verification":
            # Paragraph-wide spacing (SpaceBefore/SpaceAfter).
            return self._verify_all_paragraphs(pred_path, gold_path, spec)
        if method == "dropcap_verification":
            return self._verify_dropcap(pred_path, gold_path, spec, doc_in_path=doc_in_path)
        if idMso == "MultilevelListGallery":
            return self._com_primary_with_saved_guardrail(
                lambda: self._verify_all_paragraphs(pred_path, gold_path, spec),
                lambda: self._verify_saved_multilevel_list(pred_path, gold_path),
            )

        is_property_based_family = method in (
            "property_based",
            "xml_property_check",
            "openxml_comparison",
        )

        # Layer 2: property_based-family idMso special-cases.
        # These override select_text because the dedicated handler knows
        # better than the spec metadata what the right comparison is.
        if is_property_based_family:
            if idMso == "LigatureGalleryWord":
                return self._com_primary_with_saved_guardrail(
                    lambda: self._verify_ligatures_com(pred_path, gold_path, doc_in_path=doc_in_path),
                    lambda: self._verify_ligatures(pred_path, gold_path, spec),
                )
            if idMso == "QuickStylesGallery":
                return self._verify_quick_styles(pred_path, gold_path, spec)

        # Layer 3: select_text dispatch for everything else.
        if select_text == "all":
            return self._verify_all_paragraphs(pred_path, gold_path, spec)
        if isinstance(select_text, str) and select_text.startswith("paragraph_"):
            try:
                para_num = int(select_text.split("_")[1])
            except (ValueError, IndexError):
                pass  # Malformed `paragraph_X`: fall through to literal text search.
            else:
                return self._verify_specific_paragraph(pred_path, gold_path, para_num, spec)

        if is_property_based_family:
            return self._verify_property_based(pred_path, gold_path, spec)

        # Default: compare every formatting property on the literal target text.
        return self._verify_formatting_match(pred_path, gold_path, spec)

    def _verify_property_based(
        self, pred_path: Path, gold_path: Path, spec: Dict
    ) -> VerificationResult:
        return families_formatting_verify.verify_property_based(
            pred_path,
            gold_path,
            spec,
            read_all_paragraph_character_formatting=lambda *args, **kwargs: self._core.extract_all_paragraph_character_formatting(
                *args, **kwargs
            ),
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dropcap_properties=lambda *args, **kwargs: self._core.extract_dropcap_properties(
                *args, **kwargs
            ),
            read_target_glyph_state=lambda *args, **kwargs: self._core.extract_target_glyph_state(
                *args, **kwargs
            ),
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )

    _composition_requires_target_visibility = staticmethod(
        families_formatting_visibility.composition_requires_target_visibility
    )

    _glyph_paint_contribution = staticmethod(
        families_formatting_visibility.glyph_paint_contribution
    )

    _effective_target_glyph_visible = staticmethod(
        families_formatting_visibility.effective_target_glyph_visible
    )

    def _verify_target_glyph_visibility(
        self, pred_path: Path, gold_path: Path, spec: Dict, target_text: str
    ) -> Optional[VerificationResult]:
        return families_formatting_visibility.verify_target_glyph_visibility(
            pred_path,
            gold_path,
            spec,
            target_text,
            read_target_glyph_state=lambda *args, **kwargs: self._core.extract_target_glyph_state(
                *args, **kwargs
            ),
        )

    def _verify_control_glyph_visibility(
        self, pred_path: Path, gold_path: Path, spec: Dict, control_texts: List[str]
    ) -> Optional[VerificationResult]:
        return families_formatting_visibility.verify_control_glyph_visibility(
            pred_path,
            gold_path,
            spec,
            control_texts,
            read_target_glyph_state=lambda *args, **kwargs: self._core.extract_target_glyph_state(
                *args, **kwargs
            ),
        )

    def _verify_ligatures_com(self, pred_path: Path, gold_path: Path, *, doc_in_path=None) -> VerificationResult:
        result = families_formatting_verify.verify_ligatures_com(
            pred_path,
            gold_path,
            read_all_paragraph_font_properties=lambda *args, **kwargs: self._core.extract_all_paragraph_font_properties(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
        )
        if result.passed:
            preservation = families_formatting_preservation.check_dedicated_character_preservation(
                pred_path, gold_path, doc_in_path=doc_in_path,
                read_characters=self._core.extract_all_paragraph_character_formatting)
            if preservation is not None:
                return preservation
        return result

    def _verify_ligatures(self, pred_path: Path, gold_path: Path, spec: Dict) -> VerificationResult:
        return families_formatting_verify_saved.verify_ligatures(
            pred_path,
            gold_path,
            spec,
        )

    _extract_saved_ligature_signatures = staticmethod(
        families_formatting_extract_saved.extract_saved_ligature_signatures
    )

    _extract_saved_target_run = staticmethod(
        families_formatting_extract_saved.extract_saved_target_run
    )

    def _verify_saved_font_composition(
        self, pred_path: Path, gold_path: Path, spec: Dict
    ) -> VerificationResult:
        return families_formatting_verify_saved.verify_saved_font_composition(
            pred_path,
            gold_path,
            spec,
        )

    _extract_saved_target_text_effect = staticmethod(
        families_formatting_extract_saved.extract_saved_target_text_effect
    )

    def _verify_saved_target_text_effect(
        self, pred_path: Path, gold_path: Path, spec: Dict
    ) -> VerificationResult:
        return families_formatting_verify_saved.verify_saved_target_text_effect(
            pred_path,
            gold_path,
            spec,
        )

    _extract_saved_multilevel_list = staticmethod(
        families_formatting_extract_saved.extract_saved_multilevel_list
    )

    def _verify_saved_multilevel_list(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_formatting_verify_saved.verify_saved_multilevel_list(
            pred_path,
            gold_path,
        )

    def _verify_quick_styles(
        self, pred_path: Path, gold_path: Path, spec: Dict
    ) -> VerificationResult:
        return families_formatting_verify.verify_quick_styles(
            pred_path,
            gold_path,
            spec,
            read_all_paragraph_font_properties=lambda *args, **kwargs: self._core.extract_all_paragraph_font_properties(
                *args, **kwargs
            ),
            read_body_paragraph_styles=lambda *args, **kwargs: self._core.extract_body_paragraph_styles(
                *args, **kwargs
            ),
        )

    def _verify_quick_styles_com(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_formatting_verify.verify_quick_styles_com(
            pred_path,
            gold_path,
            read_all_paragraph_font_properties=lambda *args, **kwargs: self._core.extract_all_paragraph_font_properties(
                *args, **kwargs
            ),
            read_body_paragraph_styles=lambda *args, **kwargs: self._core.extract_body_paragraph_styles(
                *args, **kwargs
            ),
        )

    def _verify_quick_styles_saved(
        self, pred_path: Path, gold_path: Path, spec: Dict
    ) -> VerificationResult:
        return families_formatting_verify_saved.verify_quick_styles_saved(
            pred_path,
            gold_path,
            spec,
        )

    _extract_quick_style_signatures = staticmethod(
        families_formatting_extract_saved.extract_quick_style_signatures
    )

    def _verify_composition_atomic_targets(
        self, pred_path: Path, gold_path: Path, spec: Dict
    ) -> Optional[VerificationResult]:
        return families_formatting_composition.verify_composition_atomic_targets(
            pred_path,
            gold_path,
            spec,
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
        )

    _normalize_dropcap_paragraphs = staticmethod(
        families_formatting_compare.normalize_dropcap_paragraphs
    )

    def _add_composition_atomic_details(
        self,
        details: Dict[str, Any],
        pred_path: Path,
        gold_path: Path,
        spec: Dict,
        target_text: str,
    ) -> None:
        return families_formatting_composition.add_composition_atomic_details(
            details,
            pred_path,
            gold_path,
            spec,
            target_text,
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dropcap_properties=lambda *args, **kwargs: self._core.extract_dropcap_properties(
                *args, **kwargs
            ),
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )

    def _build_composition_atomic_results(
        self, pred_path: Path, gold_path: Path, spec: Dict, target_text: str
    ) -> List[Dict[str, Any]]:
        return families_formatting_composition.build_composition_atomic_results(
            pred_path,
            gold_path,
            spec,
            target_text,
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dropcap_properties=lambda *args, **kwargs: self._core.extract_dropcap_properties(
                *args, **kwargs
            ),
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )

    def _evaluate_composition_target(
        self,
        target: Dict[str, Any],
        pred_path: Path,
        gold_path: Path,
        target_text: str,
        spec: Dict,
        cache: Dict[str, Any],
        operation: Optional[str] = None,
        atomic_case: Optional[str] = None,
    ) -> Dict[str, Any]:
        return families_formatting_composition.evaluate_composition_target(
            target,
            pred_path,
            gold_path,
            target_text,
            spec,
            cache,
            operation,
            atomic_case,
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dropcap_properties=lambda *args, **kwargs: self._core.extract_dropcap_properties(
                *args, **kwargs
            ),
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )

    _character_shading_present = staticmethod(families_formatting_compare.character_shading_present)

    _character_shading_state = staticmethod(families_formatting_compare.character_shading_state)

    def _composition_target_values_match(
        self,
        prop: str,
        gold_val: Any,
        pred_val: Any,
        extractor: str,
        gold_fmt: Dict[str, Any],
        pred_fmt: Dict[str, Any],
        spec: Dict,
        comparison: Optional[str] = None,
    ) -> bool:
        return families_formatting_compare.composition_target_values_match(
            prop,
            gold_val,
            pred_val,
            extractor,
            gold_fmt,
            pred_fmt,
            spec,
            comparison,
        )

    _text_outline_dash_styles_match = staticmethod(
        families_formatting_compare.text_outline_dash_styles_match
    )

    _round_dot_text_outline_requested = staticmethod(
        families_formatting_compare.round_dot_text_outline_requested
    )

    _reference_comparison_matches = staticmethod(
        families_formatting_compare.reference_comparison_matches
    )

    _composition_indent_comparison = staticmethod(
        families_formatting_compare.composition_indent_comparison
    )

    _composition_target_comparison = staticmethod(
        families_formatting_compare.composition_target_comparison
    )

    _composition_preserves_character_formatting = staticmethod(
        families_formatting_compare.composition_preserves_character_formatting
    )

    def _check_character_formatting_preservation(
        self, pred_path: Path, gold_path: Path, spec: Dict, select_text: str
    ) -> Optional[VerificationResult]:
        return families_formatting_preservation.check_character_formatting_preservation(
            pred_path,
            gold_path,
            spec,
            select_text,
            read_all_paragraph_character_formatting=lambda *args, **kwargs: self._core.extract_all_paragraph_character_formatting(
                *args, **kwargs
            ),
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dropcap_properties=lambda *args, **kwargs: self._core.extract_dropcap_properties(
                *args, **kwargs
            ),
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )

    def _composition_uses_bullet_normalization(self, spec: Dict) -> bool:
        return families_formatting_compare.composition_uses_bullet_normalization(
            spec,
        )

    _is_alignment_visually_equivalent = staticmethod(
        families_formatting_compare.is_alignment_visually_equivalent
    )

    def _verify_formatting_match(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_formatting_verify.verify_formatting_match(
            pred_path,
            gold_path,
            spec,
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )

    def _verify_dropcap(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path=None
    ) -> VerificationResult:
        result = families_formatting_paragraphs.verify_dropcap(
            pred_path,
            gold_path,
            spec,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dropcap_properties=lambda *args, **kwargs: self._core.extract_dropcap_properties(
                *args, **kwargs
            ),
        )
        if result.passed:
            preservation = families_formatting_preservation.check_dedicated_character_preservation(
                pred_path, gold_path, doc_in_path=doc_in_path, drop_cap=True,
                read_characters=self._core.extract_all_paragraph_character_formatting)
            if preservation is not None:
                return preservation
        return result

    def _verify_all_paragraphs(
        self, pred_path: Path, gold_path: Path, spec: Dict
    ) -> VerificationResult:
        return families_formatting_paragraphs.verify_all_paragraphs(
            pred_path,
            gold_path,
            spec,
            read_all_paragraph_character_formatting=lambda *args, **kwargs: self._core.extract_all_paragraph_character_formatting(
                *args, **kwargs
            ),
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dropcap_properties=lambda *args, **kwargs: self._core.extract_dropcap_properties(
                *args, **kwargs
            ),
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )

    def _verify_specific_paragraph(
        self, pred_path: Path, gold_path: Path, paragraph_index: int, spec: Dict
    ) -> VerificationResult:
        return families_formatting_paragraphs.verify_specific_paragraph(
            pred_path,
            gold_path,
            paragraph_index,
            spec,
            read_all_paragraph_character_formatting=lambda *args, **kwargs: self._core.extract_all_paragraph_character_formatting(
                *args, **kwargs
            ),
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dropcap_properties=lambda *args, **kwargs: self._core.extract_dropcap_properties(
                *args, **kwargs
            ),
            read_paragraph_formatting=lambda *args, **kwargs: self._core.extract_paragraph_formatting(
                *args, **kwargs
            ),
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )

    def _check_content_integrity(
        self, pred_path: Path, gold_path: Path
    ) -> Optional[VerificationResult]:
        return families_formatting_preservation.check_content_integrity(
            pred_path,
            gold_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
        )

    def _check_control_text(
        self, pred_path: Path, gold_path: Path, control_texts: List[str] = None
    ) -> Optional[VerificationResult]:
        return families_formatting_preservation.check_control_text(
            pred_path,
            gold_path,
            control_texts,
            read_text_formatting=lambda *args, **kwargs: self._core.extract_text_formatting(
                *args, **kwargs
            ),
        )
