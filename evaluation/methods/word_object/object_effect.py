"""
Object effect verification for WordEval.

Handles verification for object_effect behavior across all scopes:
- core_actionable (ObjectsGroup, ObjectsAlignLeft, AlignDistributeHorizontally, etc.)
- core_parameterized (ObjectRotateGallery, PicturePositionGallery, TextWrapGallery)
- contextual_actionable, contextual_parameterized, etc.

Verification is command-driven (based on idMso), dispatching to appropriate
COM API-based verification methods:
- Alignment commands → _verify_alignment()
- Distribution commands → _verify_distribution()
- Grouping commands → _verify_grouping() / _verify_ungrouping()
- Z-order commands → _verify_z_order()
- Rotation/flip commands → _verify_rotation()
- Position commands → _verify_position()
- Wrap commands → _verify_wrap()
- Table commands → _verify_table()
- Default → _verify_shapes_match()
"""

import io
import math
import posixpath
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from .composition import (
    atomic_case_command_id,
    composition_command_ids,
    composition_steps,
)
from .core import DocumentOpenError, WordObjectVerification
from . import effect_commands

from .families.charts import (
    compare as chart_compare,
    extract_saved as chart_saved,
    verify as chart_verify,
)
from .families.charts.commands import EFFECT_COMMANDS as CHART_EFFECT_COMMANDS
from .composition import (
    composition_step_key,
    composition_step_declaration_error,
    unhandled_composition_step,
)


from .composition_visual import _get_last_step_idMso


from .families.equations import compare as families_equations_compare
from .families.equations import extract_saved as families_equations_extract_saved
from .families.equations import verify as families_equations_verify
from .families.tables import compare as families_tables_compare
from .families.tables import extract_saved as families_tables_extract_saved
from .families.tables import verify as families_tables_verify
from .preservation import structured as preservation_structured
from .preservation import visual as preservation_visual


from .families.icons import saved_color as families_icons_saved_color
from .families.icons import verify as families_icons_verify
from .families.pictures import compare as families_pictures_compare
from .families.pictures import extract_saved as families_pictures_extract_saved
from .families.pictures import verify as families_pictures_verify
from .families.shapes import effects as families_shapes_effects
from .families.shapes import extract_saved as families_shapes_extract_saved
from .families.shapes import geometry as families_shapes_geometry
from .families.shapes import layout as families_shapes_layout
from .families.shapes import lifecycle as families_shapes_lifecycle
from .families.shapes import style as families_shapes_style
from .families.shapes import text as families_shapes_text
from .families.smartart import extract_saved as families_smartart_extract_saved
from .families.smartart import refresh as families_smartart_refresh
from .families.smartart import structure as families_smartart_structure
from .families.smartart import text as families_smartart_text
from .families.smartart import verify as families_smartart_verify


from .families.metadata import effects as families_metadata_effects


from . import composition_effects as composition_effects
from . import composition_layout as composition_layout
from . import composition_visual as composition_visual
from .families.shapes import insert_bevel as families_shapes_insert_bevel
from .families.shapes import verify_collection as families_shapes_verify_collection
from .policies import object_effect as policies_object_effect
from .preservation import targets as preservation_targets


class ObjectEffectVerifier:
    """
    Verifier for object_effect behavior.

    Uses WordObjectVerification core to extract shape properties via COM API
    and compare them between predicted and gold documents.

    Dispatches to command-specific verification based on idMso.
    """

    WORDPROCESSING_DRAWING_NS = effect_commands.WORDPROCESSING_DRAWING_NS
    DRAWINGML_NS = effect_commands.DRAWINGML_NS
    WORDPROCESSING_SHAPE_NS = effect_commands.WORDPROCESSING_SHAPE_NS
    WORDPROCESSING_GROUP_NS = effect_commands.WORDPROCESSING_GROUP_NS
    WORDPROCESSINGML_NS = effect_commands.WORDPROCESSINGML_NS
    MARKUP_COMPATIBILITY_NS = effect_commands.MARKUP_COMPATIBILITY_NS
    OFFICE_MATH_NS = effect_commands.OFFICE_MATH_NS
    PICTURE_NS = effect_commands.PICTURE_NS
    EMU_PER_POINT = effect_commands.EMU_PER_POINT
    ROTATION_UNITS_PER_DEGREE = effect_commands.ROTATION_UNITS_PER_DEGREE
    WRAP_POLYGON_COORD_MAX = effect_commands.WRAP_POLYGON_COORD_MAX
    WRAP_POLYGON_POINT_TOLERANCE = effect_commands.WRAP_POLYGON_POINT_TOLERANCE

    @staticmethod
    def _picture_extraction_failure(error: Any, *, idMso: str = "") -> VerificationResult:
        return families_pictures_compare.picture_extraction_failure(
            error,
            idMso=idMso,
        )

    _extract_saved_shape_properties = staticmethod(
        families_shapes_extract_saved.extract_saved_shape_properties
    )

    _extract_saved_shape_style_signatures = staticmethod(
        families_shapes_extract_saved.extract_saved_shape_style_signatures
    )

    def _extract_shape_pair(
        self, pred_path: Path, gold_path: Path, *, prefer_saved: bool = False
    ) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]], str, Optional[str]]:
        return families_shapes_layout.extract_shape_pair(
            pred_path,
            gold_path,
            prefer_saved=prefer_saved,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    _saved_group_root_defect = staticmethod(families_shapes_extract_saved.saved_group_root_defect)

    _match_shapes_by_name = staticmethod(families_shapes_geometry.match_shapes_by_name)

    _extract_wrap_polygon_signatures = staticmethod(
        families_shapes_extract_saved.extract_wrap_polygon_signatures
    )

    _extract_saved_chart_properties = staticmethod(chart_saved.extract_properties)

    _extract_saved_smartart_connector_styles = staticmethod(
        families_smartart_extract_saved.extract_saved_smartart_connector_styles
    )

    _extract_saved_table_style_signatures = staticmethod(
        families_tables_extract_saved.extract_saved_table_style_signatures
    )

    _extract_saved_table_composition_signatures = staticmethod(
        families_tables_extract_saved.extract_saved_table_composition_signatures
    )

    _preferred_column_spread = staticmethod(families_tables_compare.preferred_column_spread)

    @staticmethod
    def _normalize_svg_color(
        value: Optional[str], *, current_color: Optional[str] = None
    ) -> Optional[str]:
        return families_icons_saved_color.normalize_svg_color(
            value,
            current_color=current_color,
        )

    _saved_svg_stroke_color_count = staticmethod(
        families_icons_saved_color.saved_svg_stroke_color_count
    )

    _extract_saved_equation_operator_signature = staticmethod(
        families_equations_extract_saved.extract_saved_equation_operator_signature
    )

    _extract_saved_logarithm_with_base_signature = staticmethod(
        families_equations_extract_saved.extract_saved_logarithm_with_base_signature
    )

    _extract_saved_three_conditions_signature = staticmethod(
        families_equations_extract_saved.extract_saved_three_conditions_signature
    )

    _extract_picture_saturation_persistence = staticmethod(
        families_pictures_extract_saved.extract_picture_saturation_persistence
    )

    _closed_polygon_core = staticmethod(families_shapes_geometry.closed_polygon_core)

    _cyclic_polygon_distance = staticmethod(families_shapes_geometry.cyclic_polygon_distance)

    _wrap_polygon_failures = staticmethod(families_shapes_geometry.wrap_polygon_failures)

    _extract_bevel_render_signatures = staticmethod(
        families_shapes_extract_saved.extract_bevel_render_signatures
    )

    _bevel_render_failures = staticmethod(families_shapes_geometry.bevel_render_failures)

    _extract_delimiter_matrix_signature = staticmethod(
        families_equations_extract_saved.extract_delimiter_matrix_signature
    )

    # Command sets - each maps to a specific verification method

    # Alignment commands - verify shapes are aligned by edge/center
    ALIGNMENT_COMMANDS = effect_commands.ALIGNMENT_COMMANDS

    # Distribution commands - verify even spacing between shapes
    DISTRIBUTION_COMMANDS = effect_commands.DISTRIBUTION_COMMANDS

    # Grouping commands - verify group structure
    GROUP_COMMANDS = effect_commands.GROUP_COMMANDS
    UNGROUP_COMMANDS = effect_commands.UNGROUP_COMMANDS

    # Z-order commands - verify z-order (layering) of shapes
    Z_ORDER_COMMANDS = effect_commands.Z_ORDER_COMMANDS

    # Rotation/flip commands - verify rotation angle and flip state
    ROTATION_COMMANDS = effect_commands.ROTATION_COMMANDS

    # Position commands - verify shape position only (excludes wrap_type)
    # Position presets set wrap_type as side effect, but only position is semantic
    POSITION_COMMANDS = effect_commands.POSITION_COMMANDS

    # Wrap commands - verify text wrap type
    WRAP_COMMANDS = effect_commands.WRAP_COMMANDS

    # Table commands - verify table structure and properties
    TABLE_COMMANDS = effect_commands.TABLE_COMMANDS

    # Structured commands whose correct gold REMOVES the structured target.
    # For these the presence-based gold preflight must be inverted: a correct
    # gold has zero targets, so the guardrail verifies target absence instead.
    DESTRUCTIVE_STRUCTURED_COMMANDS = effect_commands.DESTRUCTIVE_STRUCTURED_COMMANDS

    # Picture/shape transparency commands
    TRANSPARENCY_COMMANDS = effect_commands.TRANSPARENCY_COMMANDS

    # Picture crop commands - verify crop properties on inline pictures
    PICTURE_CROP_COMMANDS = effect_commands.PICTURE_CROP_COMMANDS

    # Picture format commands - verify brightness/contrast/size reset
    PICTURE_FORMAT_COMMANDS = effect_commands.PICTURE_FORMAT_COMMANDS

    # Alt text commands - verify alternative text on inline shapes
    ALT_TEXT_COMMANDS = effect_commands.ALT_TEXT_COMMANDS

    # DocInfo commands - insert fields/content controls into header
    DOCINFO_COMMANDS = effect_commands.DOCINFO_COMMANDS

    # Equation conversion commands - verify equation format/display state
    EQUATION_CONVERSION_COMMANDS = effect_commands.EQUATION_CONVERSION_COMMANDS

    # Picture-to-SmartArt conversion commands - verify resulting SmartArt layout
    PICTURE_SMARTART_COMMANDS = effect_commands.PICTURE_SMARTART_COMMANDS

    # Shadow commands - verify shadow properties on shapes
    SHADOW_COMMANDS = effect_commands.SHADOW_COMMANDS

    # Text alignment commands - verify text paragraph alignment inside textbox shapes
    # wdAlignParagraph values:
    # 0=wdAlignParagraphLeft, 1=wdAlignParagraphCenter, 2=wdAlignParagraphRight,
    # 3=wdAlignParagraphJustify, 4=wdAlignParagraphJustifyLow (word justify),
    # 5=wdAlignParagraphDistribute (letter justify)
    # Also includes TextDirection which changes text frame orientation
    TEXT_ALIGN_COMMANDS = effect_commands.TEXT_ALIGN_COMMANDS

    ALIGNMENT_COLLECTION_METRICS = effect_commands.ALIGNMENT_COLLECTION_METRICS
    DISTRIBUTION_COLLECTION_METRICS = effect_commands.DISTRIBUTION_COLLECTION_METRICS
    SHAPE_COLLECTION_PROPERTIES = effect_commands.SHAPE_COLLECTION_PROPERTIES

    # WordArt text effect commands - verify TextEffect properties
    WORDART_COMMANDS = effect_commands.WORDART_COMMANDS

    # 3D effect commands - verify ThreeD properties
    THREED_COMMANDS = effect_commands.THREED_COMMANDS

    # Shape fill/outline/color commands - verify fill and line color properties
    FILL_OUTLINE_COLOR_COMMANDS = effect_commands.FILL_OUTLINE_COLOR_COMMANDS

    # Shape effect commands - verify shadow/glow/soft edge/reflection/gradient
    SHAPE_EFFECT_COMMANDS = effect_commands.SHAPE_EFFECT_COMMANDS

    # Shape geometry commands - verify auto shape type, dash, arrow, bevel, 3D rotation
    SHAPE_GEOMETRY_COMMANDS = effect_commands.SHAPE_GEOMETRY_COMMANDS

    # SmartArt sub-node commands - these modify SmartArt internal nodes.
    # COM API can read sub-node line/arrow properties via Node.Shapes(i).Line.
    SMARTART_SUBNODE_COMMANDS = effect_commands.SMARTART_SUBNODE_COMMANDS

    # Picture effect commands - verify picture-specific effect properties
    PICTURE_EFFECT_COMMANDS = effect_commands.PICTURE_EFFECT_COMMANDS
    POSITIVE_PICTURE_EFFECT_COMMANDS = effect_commands.POSITIVE_PICTURE_EFFECT_COMMANDS
    POSITIVE_PICTURE_EFFECT_PARAM_TYPES = effect_commands.POSITIVE_PICTURE_EFFECT_PARAM_TYPES

    # Chart commands - verify chart properties
    CHART_COMMANDS_SET = effect_commands.CHART_COMMANDS_SET

    # Equation commands - verify equation structure
    EQUATION_COMMANDS_SET = effect_commands.EQUATION_COMMANDS_SET

    # SmartArt commands - verify SmartArt layout/style/color
    SMARTART_COMMANDS_SET = effect_commands.SMARTART_COMMANDS_SET

    # SmartArt structural commands - verify node-level changes
    SMARTART_STRUCTURAL_COMMANDS = effect_commands.SMARTART_STRUCTURAL_COMMANDS

    # Scope-changing composition steps use semantic operation names rather
    # than idMso values. Map them back to atomic commands so these steps are
    # not silently skipped by the primary composition verifier.
    SMARTART_OPERATION_COMMANDS = effect_commands.SMARTART_OPERATION_COMMANDS

    # The current SmartArt fail-closed/guardrail contract is intentionally
    # narrower than the shared text-effect dispatcher.
    SMARTART_TEXT_EFFECT_COMMANDS = effect_commands.SMARTART_TEXT_EFFECT_COMMANDS

    # Text effect commands on shapes/SmartArt - verify text effect properties
    TEXT_EFFECT_COMMANDS = effect_commands.TEXT_EFFECT_COMMANDS

    def __init__(self, core: WordObjectVerification):
        self._core = core

    @staticmethod
    def _saved_evidence_error(
        message: str, *, evidence_owner: str, exception: Optional[Exception] = None
    ) -> VerificationResult:
        return policies_object_effect.saved_evidence_error(
            message,
            evidence_owner=evidence_owner,
            exception=exception,
        )

    def verify_saved_fallback(
        self,
        pred_path: Path,
        gold_path: Path,
        spec: Optional[Dict] = None,
    ) -> Optional[VerificationResult]:
        """Return a registered saved-state verifier after a COM timeout."""
        spec = spec or {}
        composition = spec.get("composition", {}) or {}
        target_type = composition.get("target_object_type") or spec.get("precondition", {}).get(
            "type", ""
        )
        if target_type == "chart":

            class _ComUnavailable:
                @staticmethod
                def extract_chart_properties(_path):
                    raise RuntimeError("Word COM timed out")

            return ObjectEffectVerifier(_ComUnavailable())._verify_chart_composition(
                pred_path, gold_path, spec
            )

        if spec.get("idMso", "") == "TextAlignGallery":
            try:
                gold = self._extract_saved_text_vertical_anchors(gold_path)
            except Exception as exc:
                return self._saved_evidence_error(
                    f"Gold saved text-anchor extraction failed: {exc}",
                    evidence_owner="gold",
                    exception=exc,
                )
            if not gold:
                return self._saved_evidence_error(
                    "Gold saved text-anchor evidence is empty",
                    evidence_owner="gold",
                )
            try:
                pred = self._extract_saved_text_vertical_anchors(pred_path)
            except Exception as exc:
                return self._saved_evidence_error(
                    f"Prediction saved text-anchor extraction failed: {exc}",
                    evidence_owner="prediction",
                    exception=exc,
                )
            if pred is None:
                return self._saved_evidence_error(
                    "Prediction saved text-anchor extraction is unavailable",
                    evidence_owner="prediction",
                )
            return VerificationResult(
                passed=gold == pred,
                message=(
                    "PASS: Saved text vertical alignment matches"
                    if gold == pred
                    else "FAIL: Saved text vertical alignment differs"
                ),
                details={
                    "match_mode": "saved_ooxml_text_vertical_anchor",
                    "gold": gold,
                    "pred": pred,
                },
            )
        if (
            spec.get("idMso", "") == "EquationOperatorGallery"
            or composition.get("interaction_class") == "insert_delimiter_then_matrix"
        ):
            return self._verify_equation_saved_special(pred_path, gold_path, spec)
        if spec.get("idMso", "") == "BevelShapeGallery":
            try:
                gold = self._extract_bevel_render_signatures(gold_path)
            except Exception as exc:
                return self._saved_evidence_error(
                    f"Gold saved bevel extraction failed: {exc}",
                    evidence_owner="gold",
                    exception=exc,
                )
            if not gold:
                return self._saved_evidence_error(
                    "Gold saved bevel evidence is empty",
                    evidence_owner="gold",
                )
            try:
                pred = self._extract_bevel_render_signatures(pred_path)
            except Exception as exc:
                return self._saved_evidence_error(
                    f"Prediction saved bevel extraction failed: {exc}",
                    evidence_owner="prediction",
                    exception=exc,
                )
            failures = []
            if len(gold) != len(pred):
                failures.append(
                    {
                        "property": "saved_bevel_count",
                        "gold": len(gold),
                        "pred": len(pred),
                    }
                )
            else:
                for index, (gold_sig, pred_sig) in enumerate(zip(gold, pred), start=1):
                    if gold_sig.get("bevel_prst") != pred_sig.get("bevel_prst"):
                        failures.append(
                            {
                                "property": "bevel_prst",
                                "shape_index": index,
                                "gold": gold_sig.get("bevel_prst"),
                                "pred": pred_sig.get("bevel_prst"),
                            }
                        )
                    # Keep this property list identical to the primary
                    # guardrail in _bevel_render_failures: a laxer COM-timeout
                    # fallback would accept bevels the primary path rejects.
                    for prop in (
                        "scene_present",
                        "camera_present",
                        "light_rig_present",
                        "camera_prst_present",
                        "light_rig_type_present",
                        "light_rig_direction_present",
                    ):
                        if gold_sig.get(prop) and not pred_sig.get(prop):
                            failures.append(
                                {
                                    "property": f"bevel_{prop}",
                                    "shape_index": index,
                                    "gold": True,
                                    "pred": False,
                                }
                            )
            return VerificationResult(
                passed=not failures,
                message=(
                    "PASS: Saved bevel render signature matches"
                    if not failures
                    else "FAIL: Saved bevel render signature differs"
                ),
                details={
                    "match_mode": "saved_ooxml_bevel_render_scene",
                    "differences": failures,
                },
            )
        if composition.get("interaction_class") == "insert_then_bevel":
            return self._verify_insert_then_bevel_saved(pred_path, gold_path, spec)
        if spec.get("idMso", "") == "PictureEffectsGlowGallery":
            try:
                gold = self._extract_saved_picture_glows(gold_path)
            except Exception as exc:
                return self._saved_evidence_error(
                    f"Gold saved picture-glow extraction failed: {exc}",
                    evidence_owner="gold",
                    exception=exc,
                )
            if not gold:
                return self._saved_evidence_error(
                    "Gold saved picture-glow evidence is empty",
                    evidence_owner="gold",
                )
            try:
                pred = self._extract_saved_picture_glows(pred_path)
            except Exception as exc:
                return self._saved_evidence_error(
                    f"Prediction saved picture-glow extraction failed: {exc}",
                    evidence_owner="prediction",
                    exception=exc,
                )
            if pred is None:
                return self._saved_evidence_error(
                    "Prediction saved picture-glow extraction is unavailable",
                    evidence_owner="prediction",
                )
            return VerificationResult(
                passed=gold == pred,
                message=(
                    "PASS: Saved picture glow signature matches"
                    if gold == pred
                    else "FAIL: Saved picture glow signature differs"
                ),
                details={
                    "match_mode": "saved_ooxml_picture_glow_presence",
                    "gold": gold,
                    "pred": pred,
                },
            )
        if composition.get("interaction_class") == "wrap_then_position":

            class _ComUnavailable:
                POSITION_TOLERANCE_PT = 1.0

                @staticmethod
                def extract_shape_properties(_path):
                    raise RuntimeError("Word COM timed out")

            return ObjectEffectVerifier(_ComUnavailable())._verify_multi_shape_composition(
                pred_path, gold_path, spec
            )
        return None

    def _verify_insert_then_bevel_saved(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_insert_bevel.verify_insert_then_bevel_saved(
            pred_path,
            gold_path,
            spec,
            extract_bevel_render_signatures=lambda *args, **kwargs: self._extract_bevel_render_signatures(
                *args, **kwargs
            ),
            extract_saved_shape_properties=lambda *args, **kwargs: self._extract_saved_shape_properties(
                *args, **kwargs
            ),
            match_shapes_by_name=lambda *args, **kwargs: self._match_shapes_by_name(
                *args, **kwargs
            ),
        )

    @staticmethod
    def _com_primary_with_saved_guardrail(
        com_verify, saved_verify, *, saved_pass_after_com_failure: bool = False
    ):
        return policies_object_effect.com_primary_with_saved_guardrail(
            com_verify,
            saved_verify,
            saved_pass_after_com_failure=saved_pass_after_com_failure,
        )

    def verify(
        self,
        pred_path: Path,
        gold_path: Path,
        spec: Optional[Dict] = None,
        *,
        doc_in_path: Optional[Path] = None,
    ) -> VerificationResult:
        """Verify the requested object state and shared visible preservation.

        Shape/layout/text and picture/SmartArt commands are visually scoped:
        a prediction must both apply the requested object change and preserve
        visible document content and page structure outside that object.  The
        route-specific verifier remains authoritative for the target object;
        the saved package guardrail below owns only cross-route collateral
        evidence and missing-gold classification.
        """
        guardrail_kind = self._visual_object_guardrail_kind(spec)
        # Several focused unit tests intentionally exercise route semantics
        # with FakeCore data and placeholder paths.  Package guardrails apply
        # only when real packages were supplied; production evaluation always
        # reaches this verifier with existing, preflighted paths.
        if guardrail_kind and (not pred_path.exists() or not gold_path.exists()):
            guardrail_kind = None
        if guardrail_kind:
            try:
                gold_has_target = self._saved_visual_target_present(gold_path, guardrail_kind)
            except Exception as exc:
                error = f"Gold visual-object evidence could not be read: {exc}"
                return VerificationResult(
                    passed=False,
                    message=f"ERROR: {error}",
                    details={
                        "error_owner": "gold_fixture",
                        "guardrail_kind": guardrail_kind,
                        "missing_evidence": "saved_visual_object_target",
                    },
                    error=error,
                )
            if not gold_has_target:
                error = (
                    f"Gold has no {guardrail_kind} target evidence for the "
                    "requested object-effect route"
                )
                return VerificationResult(
                    passed=False,
                    message=f"ERROR: {error}",
                    details={
                        "error_owner": "gold_fixture",
                        "guardrail_kind": guardrail_kind,
                        "missing_evidence": "saved_visual_object_target",
                    },
                    error=error,
                )

        result = self._verify_requested_semantics(pred_path, gold_path, spec, doc_in_path=doc_in_path)
        if not guardrail_kind or not result.passed:
            return result

        composition_steps = self._composition_steps(spec)
        preservation = self._verify_visual_object_preservation(
            pred_path,
            gold_path,
            allow_textbox_anchor_relocation=(
                any(step.get("operation") == "insert_textbox" for step in composition_steps)
            ),
            page_info_reader=self._core.extract_page_info,
        )
        if not preservation.passed:
            return preservation
        result.details["visible_preservation"] = preservation.details
        if self._smartart_connector_fallback_contract_is_complete(spec, result.details):
            # Word does not expose the connector arrow on the targeted
            # SmartArt node reliably through the Object Model.  The composition
            # route therefore reads that one property from the saved drawing
            # cache while every other declared atomic step remains COM-backed.
            # Once the shared visible-preservation projection has also passed,
            # this is a complete target + preservation contract.  The terminal
            # evaluator still requires its independent clean-Word open probe.
            result.details["fallback_contract"] = "complete"
        return result

    def _visual_object_guardrail_kind(self, spec: Optional[Dict]) -> Optional[str]:
        return preservation_targets.visual_object_guardrail_kind(
            spec,
        )

    _saved_visual_target_present = staticmethod(preservation_targets.saved_visual_target_present)

    _saved_visual_preservation_signature = staticmethod(
        preservation_visual.saved_visual_preservation_signature
    )

    @classmethod
    def _compare_saved_visual_preservation_signatures(
        cls,
        gold: Dict[str, Any],
        pred: Dict[str, Any],
        *,
        allow_textbox_anchor_relocation: bool = False,
        allow_chart_terminal_empty: bool = False,
    ) -> Dict[str, Any]:
        return preservation_visual.compare_saved_visual_preservation_signatures(
            gold,
            pred,
            allow_textbox_anchor_relocation=allow_textbox_anchor_relocation,
            allow_chart_terminal_empty=allow_chart_terminal_empty,
        )

    def _structured_object_guardrail_kind(self, spec: Optional[Dict]) -> Optional[str]:
        return preservation_targets.structured_object_guardrail_kind(
            spec,
        )

    _saved_structured_target_present = staticmethod(
        preservation_structured.saved_structured_target_present
    )

    def _verify_structured_object_preservation(
        self,
        pred_path: Path,
        gold_path: Path,
        kind: str,
        *,
        allow_chart_terminal_empty: bool = False,
    ) -> VerificationResult:
        return preservation_structured.verify_structured_object_preservation(
            pred_path,
            gold_path,
            kind,
            allow_chart_terminal_empty=allow_chart_terminal_empty,
            read_page_info=lambda *args, **kwargs: self._core.extract_page_info(*args, **kwargs),
        )

    def _structured_gold_preflight(
        self, gold_path: Path, kind: Optional[str], *, expect_present: bool = True
    ) -> Optional[VerificationResult]:
        return preservation_structured.structured_gold_preflight(
            gold_path,
            kind,
            expect_present=expect_present,
        )

    def _finalize_structured_result(
        self,
        result: VerificationResult,
        pred_path: Path,
        gold_path: Path,
        kind: Optional[str],
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        return preservation_structured.finalize_structured_result(
            result,
            pred_path,
            gold_path,
            kind,
            spec,
            composition_steps=lambda *args, **kwargs: self._composition_steps(*args, **kwargs),
            read_page_info=lambda *args, **kwargs: self._core.extract_page_info(*args, **kwargs),
        )

    @classmethod
    def _verify_visual_object_preservation(
        cls,
        pred_path: Path,
        gold_path: Path,
        *,
        allow_textbox_anchor_relocation: bool = False,
        page_info_reader: Optional[Callable[[str], Dict[str, Any]]] = None,
    ) -> VerificationResult:
        return preservation_visual.verify_visual_object_preservation(
            pred_path,
            gold_path,
            allow_textbox_anchor_relocation=allow_textbox_anchor_relocation,
            page_info_reader=page_info_reader,
        )

    def _verify_requested_semantics(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None,
        *, doc_in_path: Optional[Path] = None,
    ) -> VerificationResult:
        """
        Verify object effect based on command type (idMso).

        Dispatches to appropriate COM API-based verification method.
        """
        if not spec:
            return self._verify_shapes_match(pred_path, gold_path)

        structured_kind = self._structured_object_guardrail_kind(spec)
        expect_target_present = spec.get("idMso") not in self.DESTRUCTIVE_STRUCTURED_COMMANDS
        preflight = self._structured_gold_preflight(
            gold_path, structured_kind, expect_present=expect_target_present
        )
        if preflight is not None:
            return preflight

        # Composition dispatch: route by target_object_type (interaction
        # compositions) or precondition_type (additive compositions).
        if spec.get("composition"):
            target_type = spec.get("composition", {}).get("target_object_type", "")
            # Fall back to precondition.type for additive compositions
            if not target_type:
                target_type = spec.get("precondition", {}).get("type", "")

            result = None
            if target_type == "table":
                result = self._verify_table(pred_path, gold_path, spec, doc_in_path=doc_in_path)
            elif target_type == "chart":
                result = self._verify_chart_composition(pred_path, gold_path, spec)
            elif target_type == "picture":
                result = self._verify_picture_composition(pred_path, gold_path, spec)
            elif target_type in (
                "smartart",
                "smartart_list",
                "smartart_hierarchy",
                "smartart_orgchart",
            ):
                result = self._verify_smartart_composition(pred_path, gold_path, spec)
            elif target_type == "equation":
                result = self._verify_equation(pred_path, gold_path, spec)
            elif target_type == "textbox":
                result = self._verify_textbox_composition(pred_path, gold_path, spec)
            elif target_type == "shape":
                # Layout-coupled single-shape compositions (wrap + position,
                # resize + wrap, etc.) must compare position/wrap_type/etc.
                # Route them through multi-shape property comparison, which
                # handles 1-shape inputs correctly.
                composition_type = spec.get("composition", {}).get("type", "")
                if composition_type == "layout_coupled":
                    result = self._verify_multi_shape_composition(pred_path, gold_path, spec)
                else:
                    result = self._verify_shape_composition(pred_path, gold_path, spec)
            elif target_type == "multi_shape":
                result = self._verify_multi_shape_composition(pred_path, gold_path, spec)
            elif target_type == "graphic":
                result = self._verify_shape_effects(pred_path, gold_path, spec)
            # Other composition types fall through to default dispatch
            if result is not None:
                if target_type in {"chart", "equation", "table"}:
                    route_result_passed = result.passed
                    result = self._finalize_structured_result(
                        result, pred_path, gold_path, structured_kind, spec
                    )
                    if route_result_passed and not result.passed:
                        return result
                if result.details.get("verification_source") in {
                    "saved_ooxml",
                    "saved_ooxml_fallback",
                    "saved_ooxml_only_wom_unavailable",
                }:
                    return result
                if (
                    target_type == "equation"
                    and spec.get("composition", {}).get("interaction_class")
                    == "insert_delimiter_then_matrix"
                ):
                    # This composition has a complete saved-OMML semantic
                    # verifier and must not fall back into COM only to add
                    # redundant intent details.
                    return result
                if (
                    target_type == "shape"
                    and spec.get("composition", {}).get("interaction_class") == "insert_then_shadow"
                ):
                    # The insert geometry/position and saved DrawingML shadow
                    # are already complete; atomic diagnostics would reopen
                    # Word only to read the same properties.
                    return result
                if (
                    target_type == "table"
                    and spec.get("composition", {}).get("interaction_class")
                    == "insert_row_column_split_then_distribute_scopes"
                ):
                    # This route validates all five operations against the
                    # input and records their native evidence. Re-comparing
                    # raw dimensions to gold would contradict valid insertion
                    # inheritance in the attached atomic diagnostics.
                    return result
                if target_type not in {
                    "smartart",
                    "smartart_list",
                    "smartart_hierarchy",
                    "smartart_orgchart",
                }:
                    self._add_composition_atomic_details(result.details, pred_path, gold_path, spec)
                return result

        idMso = spec.get("idMso", "")

        # Dispatch based on command type
        if idMso in self.ALIGNMENT_COMMANDS:
            alignment_type = self.ALIGNMENT_COMMANDS[idMso]
            return self._verify_alignment(pred_path, gold_path, alignment_type)

        if idMso in self.DISTRIBUTION_COMMANDS:
            direction = self.DISTRIBUTION_COMMANDS[idMso]
            return self._verify_distribution(pred_path, gold_path, direction)

        if idMso in self.GROUP_COMMANDS:
            return self._verify_grouping(pred_path, gold_path)

        if idMso in self.UNGROUP_COMMANDS:
            return self._verify_ungrouping(pred_path, gold_path)

        if idMso in self.Z_ORDER_COMMANDS:
            return self._verify_z_order(pred_path, gold_path)

        if idMso in self.ROTATION_COMMANDS:
            return self._verify_rotation(pred_path, gold_path, spec)

        if idMso in self.POSITION_COMMANDS:
            return self._verify_position(pred_path, gold_path, spec)

        if idMso in self.WRAP_COMMANDS:
            return self._verify_wrap(pred_path, gold_path, spec)

        # SmartArt sub-node commands must be checked BEFORE TABLE_COMMANDS
        # because some commands (e.g. OutlineWeightGallery) exist in both sets.
        if idMso in self.SMARTART_SUBNODE_COMMANDS:
            precondition_type = spec.get("precondition", {}).get("type", "") if spec else ""
            if precondition_type in ("smartart",):
                return self._verify_smartart(pred_path, gold_path, spec)
            # Non-SmartArt context: fall through to other checks (e.g. TABLE_COMMANDS)

        if idMso in self.TABLE_COMMANDS:
            return self._finalize_structured_result(
                self._verify_table(pred_path, gold_path, spec),
                pred_path,
                gold_path,
                structured_kind,
                spec,
            )

        if idMso in self.TRANSPARENCY_COMMANDS:
            # PictureTransparencyGallery can operate on inline pictures (not floating shapes).
            precondition_type = spec.get("precondition", {}).get("type", "") if spec else ""
            if precondition_type == "picture":
                return self._verify_picture_effects(pred_path, gold_path, spec)
            return self._verify_transparency(pred_path, gold_path, spec)

        if idMso in self.ALT_TEXT_COMMANDS:
            return self._verify_alt_text(pred_path, gold_path, spec)

        if idMso in self.DOCINFO_COMMANDS:
            return self._verify_docinfo(pred_path, gold_path, spec)

        if idMso in self.EQUATION_CONVERSION_COMMANDS:
            return self._verify_equation_conversion(pred_path, gold_path, spec)

        if idMso in self.PICTURE_SMARTART_COMMANDS:
            return self._verify_picture_smartart(pred_path, gold_path, spec)

        if idMso in self.PICTURE_FORMAT_COMMANDS:
            return self._verify_picture_format(pred_path, gold_path, spec)

        if idMso in self.PICTURE_CROP_COMMANDS:
            return self._verify_picture_crop(pred_path, gold_path, spec)

        if idMso in self.SHADOW_COMMANDS:
            return self._verify_shadow(pred_path, gold_path, spec)

        if idMso in self.TEXT_ALIGN_COMMANDS:
            return self._verify_text_alignment(pred_path, gold_path, spec)

        if idMso in self.WORDART_COMMANDS:
            return self._verify_wordart(pred_path, gold_path, spec)

        if idMso in self.THREED_COMMANDS:
            return self._verify_threed(pred_path, gold_path, spec)

        if idMso in self.FILL_OUTLINE_COLOR_COMMANDS:
            return self._verify_fill_outline_color(pred_path, gold_path, spec)

        if idMso in self.SHAPE_EFFECT_COMMANDS:
            return self._verify_shape_effects(pred_path, gold_path, spec)

        if idMso in self.SHAPE_GEOMETRY_COMMANDS:
            return self._verify_shape_geometry(pred_path, gold_path, spec)

        if idMso in self.PICTURE_EFFECT_COMMANDS:
            return self._verify_picture_effects(pred_path, gold_path, spec)

        if idMso in self.CHART_COMMANDS_SET:
            return self._finalize_structured_result(
                self._verify_chart(pred_path, gold_path, spec),
                pred_path,
                gold_path,
                structured_kind,
                spec,
            )

        if idMso in self.EQUATION_COMMANDS_SET:
            return self._finalize_structured_result(
                self._verify_equation(pred_path, gold_path, spec),
                pred_path,
                gold_path,
                structured_kind,
                spec,
            )

        if idMso in self.SMARTART_COMMANDS_SET:
            return self._verify_smartart(pred_path, gold_path, spec)

        if idMso in self.SMARTART_STRUCTURAL_COMMANDS:
            return self._verify_smartart_structural(pred_path, gold_path, spec)

        if idMso in self.TEXT_EFFECT_COMMANDS:
            return self._verify_text_effects(pred_path, gold_path, spec)

        # Default: compare all shape properties
        return self._verify_shapes_match(pred_path, gold_path, spec)

    def _composition_steps(self, spec: Optional[Dict]) -> List[Dict]:
        return composition_effects.composition_steps(
            spec,
        )

    def _scoped_spec(self, spec: Optional[Dict], idMso: str) -> Dict:
        return composition_effects.scoped_spec(
            spec,
            idMso,
        )

    _uses_named_picture_effect_projection = staticmethod(
        families_pictures_compare.uses_named_picture_effect_projection
    )

    _uses_named_shape_effect_projection = staticmethod(
        families_shapes_style.uses_named_shape_effect_projection
    )

    _uses_tight_wrap_polygon_projection = staticmethod(
        composition_layout.uses_tight_wrap_polygon_projection
    )

    def _merge_sub_results(self, message: str, sub_results: Dict) -> VerificationResult:
        return composition_effects.merge_sub_results(
            message,
            sub_results,
        )

    _composition_step_key = staticmethod(composition_step_key)

    _composition_step_declaration_error = staticmethod(composition_step_declaration_error)

    _unhandled_composition_step = staticmethod(unhandled_composition_step)

    def _smartart_connector_fallback_contract_is_complete(
        self, spec: Optional[Dict], details: Dict[str, Any]
    ) -> bool:
        return preservation_targets.smartart_connector_fallback_contract_is_complete(
            spec,
            details,
        )

    def _add_composition_atomic_details(
        self, details: Dict, pred_path: Path, gold_path: Path, spec: Optional[Dict]
    ) -> None:
        return composition_effects.add_composition_atomic_details(
            details,
            pred_path,
            gold_path,
            spec,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
            extract_bevel_render_signatures=lambda *args, **kwargs: self._extract_bevel_render_signatures(
                *args, **kwargs
            ),
            extract_icon_info=lambda *args, **kwargs: self._extract_icon_info(*args, **kwargs),
            extract_saved_chart_properties=lambda *args, **kwargs: self._extract_saved_chart_properties(
                *args, **kwargs
            ),
            extract_saved_shape_properties=lambda *args, **kwargs: self._extract_saved_shape_properties(
                *args, **kwargs
            ),
            extract_saved_smartart_connector_styles=lambda *args, **kwargs: self._extract_saved_smartart_connector_styles(
                *args, **kwargs
            ),
            extract_smartart_node_details=lambda *args, **kwargs: self._extract_smartart_node_details(
                *args, **kwargs
            ),
            extract_table_autofit_flags=lambda *args, **kwargs: self._extract_table_autofit_flags(
                *args, **kwargs
            ),
            match_shapes_by_name=lambda *args, **kwargs: self._match_shapes_by_name(
                *args, **kwargs
            ),
            normalized_equation_text=lambda *args, **kwargs: self._normalized_equation_text(
                *args, **kwargs
            ),
            get_position_tolerance=lambda: self._core.POSITION_TOLERANCE_PT,
            read_chart_properties=lambda *args, **kwargs: self._core.extract_chart_properties(
                *args, **kwargs
            ),
            read_equation_structure=lambda *args, **kwargs: self._core.extract_equation_structure(
                *args, **kwargs
            ),
            read_picture_crop_properties=lambda *args, **kwargs: self._core.extract_picture_crop_properties(
                *args, **kwargs
            ),
            read_picture_effect_properties=lambda *args, **kwargs: self._core.extract_picture_effect_properties(
                *args, **kwargs
            ),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            read_smartart_placement=lambda *args, **kwargs: self._core.extract_smartart_placement(
                *args, **kwargs
            ),
            read_smartart_properties=lambda *args, **kwargs: self._core.extract_smartart_properties(
                *args, **kwargs
            ),
            read_smartart_subnode_properties=lambda *args, **kwargs: self._core.extract_smartart_subnode_properties(
                *args, **kwargs
            ),
            read_table_properties=lambda *args, **kwargs: self._core.extract_table_properties(
                *args, **kwargs
            ),
        )

    def _build_composition_atomic_results(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict]
    ) -> List[Dict[str, Any]]:
        return composition_effects.build_composition_atomic_results(
            pred_path,
            gold_path,
            spec,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
            extract_bevel_render_signatures=lambda *args, **kwargs: self._extract_bevel_render_signatures(
                *args, **kwargs
            ),
            extract_icon_info=lambda *args, **kwargs: self._extract_icon_info(*args, **kwargs),
            extract_saved_chart_properties=lambda *args, **kwargs: self._extract_saved_chart_properties(
                *args, **kwargs
            ),
            extract_saved_shape_properties=lambda *args, **kwargs: self._extract_saved_shape_properties(
                *args, **kwargs
            ),
            extract_saved_smartart_connector_styles=lambda *args, **kwargs: self._extract_saved_smartart_connector_styles(
                *args, **kwargs
            ),
            extract_smartart_node_details=lambda *args, **kwargs: self._extract_smartart_node_details(
                *args, **kwargs
            ),
            extract_table_autofit_flags=lambda *args, **kwargs: self._extract_table_autofit_flags(
                *args, **kwargs
            ),
            match_shapes_by_name=lambda *args, **kwargs: self._match_shapes_by_name(
                *args, **kwargs
            ),
            normalized_equation_text=lambda *args, **kwargs: self._normalized_equation_text(
                *args, **kwargs
            ),
            get_position_tolerance=lambda: self._core.POSITION_TOLERANCE_PT,
            read_chart_properties=lambda *args, **kwargs: self._core.extract_chart_properties(
                *args, **kwargs
            ),
            read_equation_structure=lambda *args, **kwargs: self._core.extract_equation_structure(
                *args, **kwargs
            ),
            read_picture_crop_properties=lambda *args, **kwargs: self._core.extract_picture_crop_properties(
                *args, **kwargs
            ),
            read_picture_effect_properties=lambda *args, **kwargs: self._core.extract_picture_effect_properties(
                *args, **kwargs
            ),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            read_smartart_placement=lambda *args, **kwargs: self._core.extract_smartart_placement(
                *args, **kwargs
            ),
            read_smartart_properties=lambda *args, **kwargs: self._core.extract_smartart_properties(
                *args, **kwargs
            ),
            read_smartart_subnode_properties=lambda *args, **kwargs: self._core.extract_smartart_subnode_properties(
                *args, **kwargs
            ),
            read_table_properties=lambda *args, **kwargs: self._core.extract_table_properties(
                *args, **kwargs
            ),
        )

    def _horizontal_center_endpoints(self, shapes: List[Dict[str, Any]]) -> List[float]:
        return composition_effects.horizontal_center_endpoints(
            shapes,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
        )

    def _evaluate_composition_target(
        self, target: Dict[str, Any], pred_path: Path, gold_path: Path, cache: Dict[str, Any]
    ) -> Dict[str, Any]:
        return composition_effects.evaluate_composition_target(
            target,
            pred_path,
            gold_path,
            cache,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
            extract_bevel_render_signatures=lambda *args, **kwargs: self._extract_bevel_render_signatures(
                *args, **kwargs
            ),
            extract_icon_info=lambda *args, **kwargs: self._extract_icon_info(*args, **kwargs),
            extract_saved_chart_properties=lambda *args, **kwargs: self._extract_saved_chart_properties(
                *args, **kwargs
            ),
            extract_saved_shape_properties=lambda *args, **kwargs: self._extract_saved_shape_properties(
                *args, **kwargs
            ),
            extract_saved_smartart_connector_styles=lambda *args, **kwargs: self._extract_saved_smartart_connector_styles(
                *args, **kwargs
            ),
            extract_smartart_node_details=lambda *args, **kwargs: self._extract_smartart_node_details(
                *args, **kwargs
            ),
            extract_table_autofit_flags=lambda *args, **kwargs: self._extract_table_autofit_flags(
                *args, **kwargs
            ),
            match_shapes_by_name=lambda *args, **kwargs: self._match_shapes_by_name(
                *args, **kwargs
            ),
            normalized_equation_text=lambda *args, **kwargs: self._normalized_equation_text(
                *args, **kwargs
            ),
            read_chart_properties=lambda *args, **kwargs: self._core.extract_chart_properties(
                *args, **kwargs
            ),
            read_equation_structure=lambda *args, **kwargs: self._core.extract_equation_structure(
                *args, **kwargs
            ),
            read_picture_crop_properties=lambda *args, **kwargs: self._core.extract_picture_crop_properties(
                *args, **kwargs
            ),
            read_picture_effect_properties=lambda *args, **kwargs: self._core.extract_picture_effect_properties(
                *args, **kwargs
            ),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            read_smartart_placement=lambda *args, **kwargs: self._core.extract_smartart_placement(
                *args, **kwargs
            ),
            read_smartart_properties=lambda *args, **kwargs: self._core.extract_smartart_properties(
                *args, **kwargs
            ),
            read_smartart_subnode_properties=lambda *args, **kwargs: self._core.extract_smartart_subnode_properties(
                *args, **kwargs
            ),
            read_table_properties=lambda *args, **kwargs: self._core.extract_table_properties(
                *args, **kwargs
            ),
        )

    def _extract_composition_target_data(
        self, extractor: str, path: Path, cache: Dict[str, Any], side: str
    ) -> Any:
        return composition_effects.extract_composition_target_data(
            extractor,
            path,
            cache,
            side,
            extract_bevel_render_signatures=lambda *args, **kwargs: self._extract_bevel_render_signatures(
                *args, **kwargs
            ),
            extract_icon_info=lambda *args, **kwargs: self._extract_icon_info(*args, **kwargs),
            extract_saved_chart_properties=lambda *args, **kwargs: self._extract_saved_chart_properties(
                *args, **kwargs
            ),
            extract_saved_shape_properties=lambda *args, **kwargs: self._extract_saved_shape_properties(
                *args, **kwargs
            ),
            extract_saved_smartart_connector_styles=lambda *args, **kwargs: self._extract_saved_smartart_connector_styles(
                *args, **kwargs
            ),
            extract_smartart_node_details=lambda *args, **kwargs: self._extract_smartart_node_details(
                *args, **kwargs
            ),
            extract_table_autofit_flags=lambda *args, **kwargs: self._extract_table_autofit_flags(
                *args, **kwargs
            ),
            read_chart_properties=lambda *args, **kwargs: self._core.extract_chart_properties(
                *args, **kwargs
            ),
            read_equation_structure=lambda *args, **kwargs: self._core.extract_equation_structure(
                *args, **kwargs
            ),
            read_picture_crop_properties=lambda *args, **kwargs: self._core.extract_picture_crop_properties(
                *args, **kwargs
            ),
            read_picture_effect_properties=lambda *args, **kwargs: self._core.extract_picture_effect_properties(
                *args, **kwargs
            ),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            read_smartart_placement=lambda *args, **kwargs: self._core.extract_smartart_placement(
                *args, **kwargs
            ),
            read_smartart_properties=lambda *args, **kwargs: self._core.extract_smartart_properties(
                *args, **kwargs
            ),
            read_smartart_subnode_properties=lambda *args, **kwargs: self._core.extract_smartart_subnode_properties(
                *args, **kwargs
            ),
            read_table_properties=lambda *args, **kwargs: self._core.extract_table_properties(
                *args, **kwargs
            ),
        )

    def _extract_table_autofit_flags(self, path: Path) -> List[Dict[str, Any]]:
        return families_tables_extract_saved.extract_table_autofit_flags(
            path,
        )

    def _composition_target_value(
        self, data: Any, prop: str, target: Optional[Dict[str, Any]] = None
    ) -> Any:
        return composition_effects.composition_target_value(
            data,
            prop,
            target,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
        )

    def _custom_composition_collection_value(self, data: List[Dict[str, Any]], prop: str) -> float:
        return composition_effects.custom_composition_collection_value(
            data,
            prop,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
        )

    def _composition_visual_boxes(self, data: List[Dict[str, Any]]) -> List[Dict[str, float]]:
        return composition_effects.composition_visual_boxes(
            data,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
        )

    def _alignment_anchor(self, shapes: List[Dict[str, Any]], metric: str) -> Optional[float]:
        return composition_effects.alignment_anchor(
            shapes,
            metric,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
        )

    def _distribution_center_endpoints(
        self, shapes: List[Dict[str, Any]], axis: str
    ) -> List[float]:
        return composition_effects.distribution_center_endpoints(
            shapes,
            axis,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
        )

    def _picture_effect_params(self, data: Dict[str, Any], effect_type: Optional[int]) -> Any:
        return composition_effects.picture_effect_params(
            data,
            effect_type,
        )

    def _picture_effect_present(self, data: Dict[str, Any], effect_type: Optional[int]) -> bool:
        return composition_effects.picture_effect_present(
            data,
            effect_type,
        )

    def _custom_composition_target_value(self, data: Dict[str, Any], prop: str) -> Any:
        return composition_effects.custom_composition_target_value(
            data,
            prop,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
        )

    def _composition_values_match(
        self, gold_value: Any, pred_value: Any, tolerance: Optional[float] = None
    ) -> bool:
        return composition_effects.composition_values_match(
            gold_value,
            pred_value,
            tolerance,
        )

    @classmethod
    def _normalize_composition_equation_value(cls, value: Any) -> Any:
        return composition_effects.normalize_composition_equation_value(
            value,
            normalized_equation_text=lambda *args, **kwargs: cls._normalized_equation_text(
                *args, **kwargs
            ),
        )

    def _verify_insert_shape_operation(
        self,
        pred_path: Path,
        gold_path: Path,
        params: Optional[Dict] = None,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        return families_shapes_lifecycle.verify_insert_shape_operation(
            pred_path,
            gold_path,
            params,
            spec,
            read_shape_placement=lambda *args, **kwargs: self._core.extract_shape_placement(
                *args, **kwargs
            ),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    def _extract_icon_info(self, docx_path: Path) -> Dict[str, Any]:
        return families_icons_verify.extract_icon_info(
            docx_path,
            read_svg_icon_properties=lambda *args, **kwargs: self._core.extract_svg_icon_properties(
                *args, **kwargs
            ),
        )

    def _verify_insert_icon_operation(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_icons_verify.verify_insert_icon_operation(
            pred_path,
            gold_path,
            read_svg_icon_properties=lambda *args, **kwargs: self._core.extract_svg_icon_properties(
                *args, **kwargs
            ),
        )

    def _verify_shape_composition(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return composition_visual.verify_shape_composition(
            pred_path,
            gold_path,
            spec,
            composition_step_declaration_error=lambda *args, **kwargs: self._composition_step_declaration_error(
                *args, **kwargs
            ),
            composition_step_key=lambda *args, **kwargs: self._composition_step_key(
                *args, **kwargs
            ),
            unhandled_composition_step=lambda *args, **kwargs: self._unhandled_composition_step(
                *args, **kwargs
            ),
            verify_fill_outline_color=lambda *args, **kwargs: self._verify_fill_outline_color(
                *args, **kwargs
            ),
            verify_insert_shape_operation=lambda *args, **kwargs: self._verify_insert_shape_operation(
                *args, **kwargs
            ),
            verify_shape_effects=lambda *args, **kwargs: self._verify_shape_effects(
                *args, **kwargs
            ),
            verify_shape_geometry=lambda *args, **kwargs: self._verify_shape_geometry(
                *args, **kwargs
            ),
            verify_text_effects=lambda *args, **kwargs: self._verify_text_effects(*args, **kwargs),
        )

    def _verify_picture_composition(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return composition_visual.verify_picture_composition(
            pred_path,
            gold_path,
            spec,
            composition_step_declaration_error=lambda *args, **kwargs: self._composition_step_declaration_error(
                *args, **kwargs
            ),
            composition_step_key=lambda *args, **kwargs: self._composition_step_key(
                *args, **kwargs
            ),
            unhandled_composition_step=lambda *args, **kwargs: self._unhandled_composition_step(
                *args, **kwargs
            ),
            verify_insert_icon_operation=lambda *args, **kwargs: self._verify_insert_icon_operation(
                *args, **kwargs
            ),
            verify_picture_crop=lambda *args, **kwargs: self._verify_picture_crop(*args, **kwargs),
            verify_picture_effects=lambda *args, **kwargs: self._verify_picture_effects(
                *args, **kwargs
            ),
            verify_picture_format=lambda *args, **kwargs: self._verify_picture_format(
                *args, **kwargs
            ),
        )

    def _verify_textbox_composition(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return composition_visual.verify_textbox_composition(
            pred_path,
            gold_path,
            spec,
            composition_step_declaration_error=lambda *args, **kwargs: self._composition_step_declaration_error(
                *args, **kwargs
            ),
            composition_step_key=lambda *args, **kwargs: self._composition_step_key(
                *args, **kwargs
            ),
            unhandled_composition_step=lambda *args, **kwargs: self._unhandled_composition_step(
                *args, **kwargs
            ),
            verify_fill_outline_color=lambda *args, **kwargs: self._verify_fill_outline_color(
                *args, **kwargs
            ),
            verify_text_alignment=lambda *args, **kwargs: self._verify_text_alignment(
                *args, **kwargs
            ),
            verify_textbox_insert=lambda *args, **kwargs: self._verify_textbox_insert(
                *args, **kwargs
            ),
            verify_textbox_text_fill=lambda *args, **kwargs: self._verify_textbox_text_fill(
                *args, **kwargs
            ),
        )

    def _verify_smartart_composition(
        self,
        pred_path: Path,
        gold_path: Path,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        sub_results = {}
        declaration_error = self._composition_step_declaration_error("smartart", spec)
        if declaration_error is not None:
            return declaration_error
        steps = self._composition_steps(spec)

        # Multiple SmartArt steps and their atomic-result diagnostics observe
        # the same immutable gold/pred files. Cache those observations only
        # for this one composition call: every step and target still executes
        # its own comparison, without reopening Word for identical extraction.
        cache_sentinel = object()
        previous_refresh_cache = getattr(self, "_smartart_refresh_projection_cache", cache_sentinel)
        previous_node_cache = getattr(self, "_smartart_node_details_cache", cache_sentinel)
        self._smartart_refresh_projection_cache = {}
        self._smartart_node_details_cache = {}
        try:
            composition_result = None
            seen_steps: set[str] = set()
            if not steps:
                composition_result = self._verify_smartart_structural(pred_path, gold_path, spec)
            for step_index, step in enumerate(steps, start=1):
                step_key, step_error = self._composition_step_key(
                    "smartart", step_index, step, seen_steps
                )
                if step_error is not None:
                    composition_result = step_error
                    break
                op = step.get("operation")
                idMso = step.get("idMso")
                structural_id = self.SMARTART_OPERATION_COMMANDS.get(op)
                if op == "insert_smartart":
                    result = self._verify_smartart_structural(pred_path, gold_path, spec)
                elif structural_id:
                    result = self._verify_smartart_structural(
                        pred_path, gold_path, self._scoped_spec(spec, structural_id)
                    )
                elif idMso in self.SMARTART_SUBNODE_COMMANDS:
                    scoped = self._scoped_spec(spec, idMso)
                    scoped["precondition"] = {"type": "smartart"}
                    result = self._verify_smartart(pred_path, gold_path, scoped)
                elif idMso in self.SMARTART_COMMANDS_SET:
                    result = self._verify_smartart(
                        pred_path, gold_path, self._scoped_spec(spec, idMso)
                    )
                elif idMso in self.SMARTART_STRUCTURAL_COMMANDS:
                    result = self._verify_smartart_structural(
                        pred_path, gold_path, self._scoped_spec(spec, idMso)
                    )
                elif idMso in self.TEXT_EFFECT_COMMANDS:
                    result = self._verify_text_effects(
                        pred_path, gold_path, self._scoped_spec(spec, idMso)
                    )
                else:
                    composition_result = self._unhandled_composition_step(
                        "smartart", step_index, step
                    )
                    break
                if not result.passed:
                    composition_result = result
                    break
                sub_results[step_key] = result.details

            if composition_result is None:
                composition_result = self._merge_sub_results(
                    "PASS: SmartArt composition properties match", sub_results
                )
            self._add_composition_atomic_details(
                composition_result.details, pred_path, gold_path, spec
            )
            return composition_result
        finally:
            if previous_refresh_cache is cache_sentinel:
                delattr(self, "_smartart_refresh_projection_cache")
            else:
                self._smartart_refresh_projection_cache = previous_refresh_cache
            if previous_node_cache is cache_sentinel:
                delattr(self, "_smartart_node_details_cache")
            else:
                self._smartart_node_details_cache = previous_node_cache

    # Alignment Verification

    _relative_layout_preservation_failures = staticmethod(
        families_shapes_geometry.relative_layout_preservation_failures
    )

    def _normalized_shape_start(self, shape: Dict[str, Any], axis: str) -> float:
        return families_shapes_geometry.normalized_shape_start(
            shape,
            axis,
        )

    def _normalized_alignment_value(self, shape: Dict[str, Any], alignment_type: str) -> float:
        return families_shapes_geometry.normalized_alignment_value(
            shape,
            alignment_type,
        )

    def _orthogonal_alignment_preservation_failures(
        self,
        pred_shapes: List[Dict[str, Any]],
        gold_shapes: List[Dict[str, Any]],
        alignment_type: str,
        tolerance: float,
    ) -> List[str]:
        return families_shapes_geometry.orthogonal_alignment_preservation_failures(
            pred_shapes,
            gold_shapes,
            alignment_type,
            tolerance,
        )

    def _verify_alignment(
        self, pred_path: Path, gold_path: Path, alignment_type: str
    ) -> VerificationResult:
        return families_shapes_layout.verify_alignment(
            pred_path,
            gold_path,
            alignment_type,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Distribution Verification

    def _verify_distribution(
        self, pred_path: Path, gold_path: Path, direction: str
    ) -> VerificationResult:
        return families_shapes_layout.verify_distribution(
            pred_path,
            gold_path,
            direction,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Grouping Verification

    def _verify_grouping(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_lifecycle.verify_grouping(
            pred_path,
            gold_path,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    def _verify_ungrouping(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_lifecycle.verify_ungrouping(
            pred_path,
            gold_path,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Z-Order Verification

    def _verify_z_order(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_lifecycle.verify_z_order(
            pred_path,
            gold_path,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Rotation/Flip Verification (for ObjectRotateGallery)

    def _verify_rotation(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_lifecycle.verify_rotation(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Position Verification (for PicturePositionGallery)

    # Threshold for detecting Word's preset position constants
    # Word uses large negative values (around -999xxx) for relative positioning
    PRESET_POSITION_THRESHOLD = effect_commands.PRESET_POSITION_THRESHOLD

    # Preset constant to position type mapping
    # These are Word's internal values for preset positions
    PRESET_VALUE_MAP = effect_commands.PRESET_VALUE_MAP

    def _is_preset_position(self, value: float) -> bool:
        return families_shapes_geometry.is_preset_position(
            value,
        )

    def _calculate_actual_position(
        self,
        raw_pos: float,
        shape_size: float,
        rel_position: int,
        page_size: float,
        margin_start: float,
        margin_end: float,
    ) -> float:
        return families_shapes_geometry.calculate_actual_position(
            raw_pos,
            shape_size,
            rel_position,
            page_size,
            margin_start,
            margin_end,
        )

    def _verify_position(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_layout.verify_position(
            pred_path,
            gold_path,
            spec,
            get_position_tolerance=lambda: self._core.POSITION_TOLERANCE_PT,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Text Wrap Verification (for TextWrapGallery)

    def _verify_wrap(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_layout.verify_wrap(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            uses_tight_wrap_polygon_projection=lambda *args, **kwargs: self._uses_tight_wrap_polygon_projection(
                *args, **kwargs
            ),
        )

    # Transparency Verification

    def _verify_transparency(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_effects.verify_transparency(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Alt Text Verification

    def _verify_alt_text(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_pictures_verify.verify_alt_text(
            pred_path,
            gold_path,
            spec,
            read_inline_shapes=lambda *args, **kwargs: self._core.extract_inline_shapes(
                *args, **kwargs
            ),
        )

    def _verify_docinfo(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_metadata_effects.verify_docinfo(
            pred_path,
            gold_path,
            spec,
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
        )

    def _verify_equation_conversion(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_equations_verify.verify_equation_conversion(
            pred_path,
            gold_path,
            spec,
            read_equation_structure=lambda *args, **kwargs: self._core.extract_equation_structure(
                *args, **kwargs
            ),
        )

    def _verify_picture_smartart(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_pictures_verify.verify_picture_smartart(
            pred_path,
            gold_path,
            spec,
            get_position_tolerance=lambda: self._core.POSITION_TOLERANCE_PT,
            read_inline_shapes=lambda *args, **kwargs: self._core.extract_inline_shapes(
                *args, **kwargs
            ),
        )

    # Picture Crop Verification

    def _verify_picture_crop(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_pictures_verify.verify_picture_crop(
            pred_path,
            gold_path,
            spec,
            read_picture_crop_properties=lambda *args, **kwargs: self._core.extract_picture_crop_properties(
                *args, **kwargs
            ),
        )

    def _verify_picture_format(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_pictures_verify.verify_picture_format(
            pred_path,
            gold_path,
            spec,
            read_picture_format_properties=lambda *args, **kwargs: self._core.extract_picture_format_properties(
                *args, **kwargs
            ),
        )

    # Shadow Verification

    def _verify_shadow(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_effects.verify_shadow(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Text Alignment Verification

    def _textbox_shapes(self, docx_path: Path) -> List[Dict]:
        return families_shapes_text.textbox_shapes(
            docx_path,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    def _verify_textbox_insert(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_text.verify_textbox_insert(
            pred_path,
            gold_path,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    def _verify_textbox_text_fill(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_text.verify_textbox_text_fill(
            pred_path,
            gold_path,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    def _verify_text_alignment(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_text.verify_text_alignment(
            pred_path,
            gold_path,
            spec,
            get_composition_steps=lambda *args, **kwargs: self._composition_steps(*args, **kwargs),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # WordArt Verification

    def _verify_wordart(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_text.verify_wordart(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # 3D Effect Verification

    def _verify_threed(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_effects.verify_threed(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Fill/Outline/Color Verification

    def _verify_fill_outline_color(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_effects.verify_fill_outline_color(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            saved_evidence_error=lambda *args, **kwargs: self._saved_evidence_error(
                *args, **kwargs
            ),
        )

    # Shape Effects Verification

    def _verify_shape_effects(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_style.verify_shape_effects(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Shape Geometry Verification

    def _verify_shape_geometry(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_style.verify_shape_geometry(
            pred_path,
            gold_path,
            spec,
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Picture Effects Verification

    _extract_saved_picture_glows = staticmethod(
        families_pictures_extract_saved.extract_saved_picture_glows
    )

    _extract_saved_glow_signatures = staticmethod(
        families_shapes_extract_saved.extract_saved_glow_signatures
    )

    _extract_saved_text_vertical_anchors = staticmethod(
        families_shapes_extract_saved.extract_saved_text_vertical_anchors
    )

    _extract_saved_picture_named_effects = staticmethod(
        families_pictures_extract_saved.extract_saved_picture_named_effects
    )

    _extract_saved_shape_named_effects = staticmethod(
        families_shapes_extract_saved.extract_saved_shape_named_effects
    )

    _outer_shadow_has_centered_zero_offset = staticmethod(
        families_pictures_compare.outer_shadow_has_centered_zero_offset
    )

    def _verify_named_shape_effect_state(
        self,
        pred_path: Path,
        gold_path: Path,
        id_mso: str,
        pred_shapes: List[Dict[str, Any]],
        gold_shapes: List[Dict[str, Any]],
    ) -> VerificationResult:
        return families_shapes_style.verify_named_shape_effect_state(
            pred_path,
            gold_path,
            id_mso,
            pred_shapes,
            gold_shapes,
        )

    def _verify_named_picture_effect_state(
        self,
        pred_path: Path,
        gold_path: Path,
        id_mso: str,
        pred_props: Dict[str, Any],
        gold_props: Dict[str, Any],
    ) -> VerificationResult:
        return families_pictures_verify.verify_named_picture_effect_state(
            pred_path,
            gold_path,
            id_mso,
            pred_props,
            gold_props,
        )

    def _verify_picture_effects(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_pictures_verify.verify_picture_effects(
            pred_path,
            gold_path,
            spec,
            read_picture_effect_properties=lambda *args, **kwargs: self._core.extract_picture_effect_properties(
                *args, **kwargs
            ),
            saved_evidence_error=lambda *args, **kwargs: self._saved_evidence_error(
                *args, **kwargs
            ),
        )

    def _verify_picture_saturation_persistence(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_pictures_verify.verify_picture_saturation_persistence(
            pred_path,
            gold_path,
            saved_evidence_error=lambda *args, **kwargs: self._saved_evidence_error(
                *args, **kwargs
            ),
        )

    def _verify_picture_effects_param(
        self,
        gold_props: Dict,
        pred_props: Dict,
        idMso: str,
        effect_type: int = None,
        compare_effect_type: bool = False,
    ) -> VerificationResult:
        return families_pictures_compare.verify_picture_effects_param(
            gold_props,
            pred_props,
            idMso,
            effect_type,
            compare_effect_type,
        )

    # Chart Verification

    _chart_diffs_for_idmso = staticmethod(chart_compare.diff_command)

    def _verify_chart(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return chart_verify.verify_effect(
            pred_path,
            gold_path,
            spec,
            read_com=lambda path: self._core.extract_chart_properties(path),
        )

    # Map composition step "operation" names to a list of properties to check.
    # Operations are higher-level than idMso (e.g. "insert_chart" with explicit
    # data params, "change_chart_type" with target type, "add_drop_lines" boolean).
    # See N1/N4 in the composition review findings.
    #
    # All comparisons are pred-vs-gold (not pred-vs-spec). Gold is the
    # authoritative ground truth — comparing against spec params instead would
    # spuriously flag properties where the extractor returns None for both
    # gold and pred (e.g. series.name on COM-extracted charts).
    _chart_diffs_for_operation = staticmethod(chart_compare.diff_operation)

    def _verify_chart_composition(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return chart_verify.verify_composition(
            pred_path,
            gold_path,
            spec,
            read_com=lambda path: self._core.extract_chart_properties(path),
        )

    # Equation Verification

    _normalized_equation_text = staticmethod(families_equations_compare.normalized_equation_text)

    _folded_equation_text = staticmethod(families_equations_compare.folded_equation_text)

    _normalized_basic_operator_text = staticmethod(
        families_equations_compare.normalized_basic_operator_text
    )

    _canonical_angle_bracket = staticmethod(families_equations_compare.canonical_angle_bracket)

    _normalized_angle_delimiter_text = staticmethod(
        families_equations_compare.normalized_angle_delimiter_text
    )

    _has_native_angle_delimiter = staticmethod(
        families_equations_compare.has_native_angle_delimiter
    )

    _equation_com_semantic_signature = staticmethod(
        families_equations_compare.equation_com_semantic_signature
    )

    def _verify_equation_com_special(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_equations_verify.verify_equation_com_special(
            pred_path,
            gold_path,
            read_equation_structure=lambda *args, **kwargs: self._core.extract_equation_structure(
                *args, **kwargs
            ),
        )

    def _verify_equation_saved_special(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_equations_verify.verify_equation_saved_special(
            pred_path,
            gold_path,
            spec,
            saved_evidence_error=lambda *args, **kwargs: self._saved_evidence_error(
                *args, **kwargs
            ),
        )

    def _verify_equation(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_equations_verify.verify_equation(
            pred_path,
            gold_path,
            spec,
            read_equation_structure=lambda *args, **kwargs: self._core.extract_equation_structure(
                *args, **kwargs
            ),
            saved_evidence_error=lambda *args, **kwargs: self._saved_evidence_error(
                *args, **kwargs
            ),
        )

    # SmartArt Verification

    def _extract_refreshed_smartart_projection(
        self,
        path: Path,
    ) -> Dict[str, Any]:
        """Return a composition-local Word-refreshed SmartArt projection."""

        refresh_cache = getattr(self, "_smartart_refresh_projection_cache", None)
        cache_key = str(path.resolve())
        if isinstance(refresh_cache, dict) and cache_key in refresh_cache:
            return refresh_cache[cache_key]

        projection = self._core.extract_smartart_visible_projection(str(path))
        if isinstance(refresh_cache, dict):
            refresh_cache[cache_key] = projection
        return projection

    def _extract_smartart_node_details(
        self,
        path: Path,
    ) -> Dict[str, Any]:
        """Return composition-local live SmartArt node details."""

        node_cache = getattr(self, "_smartart_node_details_cache", None)
        cache_key = str(path.resolve())
        if isinstance(node_cache, dict) and cache_key in node_cache:
            return node_cache[cache_key]

        details = self._core.extract_smartart_node_details(str(path))
        if isinstance(node_cache, dict):
            node_cache[cache_key] = details
        return details

    def _verify_refreshed_smartart_projection(
        self,
        pred_path: Path,
        gold_path: Path,
        gold_semantic: Dict[str, Any],
        pred_semantic: Dict[str, Any],
        raw_cache_trigger: Dict[str, Any],
        idMso: str,
    ) -> VerificationResult:
        return families_smartart_refresh.verify_refreshed_smartart_projection(
            pred_path,
            gold_path,
            gold_semantic,
            pred_semantic,
            raw_cache_trigger,
            idMso,
            extract_refreshed_smartart_projection=lambda *args, **kwargs: self._extract_refreshed_smartart_projection(
                *args, **kwargs
            ),
            get_cleanup_word=lambda: getattr(self._core, "_cleanup_word", None),
        )

    def _verify_refreshed_smartart_gallery_projection(
        self,
        pred_path: Path,
        gold_path: Path,
        gold_semantic: Dict[str, Any],
        pred_semantic: Dict[str, Any],
        raw_cache_trigger: Dict[str, Any],
        idMso: str,
    ) -> VerificationResult:
        return families_smartart_refresh.verify_refreshed_smartart_gallery_projection(
            pred_path,
            gold_path,
            gold_semantic,
            pred_semantic,
            raw_cache_trigger,
            idMso,
            extract_refreshed_smartart_projection=lambda *args, **kwargs: self._extract_refreshed_smartart_projection(
                *args, **kwargs
            ),
            get_cleanup_word=lambda: getattr(self._core, "_cleanup_word", None),
        )

    def _verify_smartart(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_smartart_verify.verify_smartart(
            pred_path,
            gold_path,
            spec,
            extract_refreshed_smartart_projection=lambda *args, **kwargs: self._extract_refreshed_smartart_projection(
                *args, **kwargs
            ),
            get_cleanup_word=lambda: getattr(self._core, "_cleanup_word", None),
            read_smartart_drawing_info=lambda *args, **kwargs: self._core.extract_smartart_drawing_info(
                *args, **kwargs
            ),
            read_smartart_properties=lambda *args, **kwargs: self._core.extract_smartart_properties(
                *args, **kwargs
            ),
            read_smartart_subnode_properties=lambda *args, **kwargs: self._core.extract_smartart_subnode_properties(
                *args, **kwargs
            ),
        )

    def _verify_smartart_structural(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_smartart_structure.verify_smartart_structural(
            pred_path,
            gold_path,
            spec,
            composition_steps=lambda *args, **kwargs: self._composition_steps(*args, **kwargs),
            extract_refreshed_smartart_projection=lambda *args, **kwargs: self._extract_refreshed_smartart_projection(
                *args, **kwargs
            ),
            extract_smartart_node_details=lambda *args, **kwargs: self._extract_smartart_node_details(
                *args, **kwargs
            ),
            get_cleanup_word=lambda: getattr(self._core, "_cleanup_word", None),
            read_smartart_drawing_info=lambda *args, **kwargs: self._core.extract_smartart_drawing_info(
                *args, **kwargs
            ),
            read_smartart_node_scale_info=lambda *args, **kwargs: self._core.extract_smartart_node_scale_info(
                *args, **kwargs
            ),
            read_smartart_placement=lambda *args, **kwargs: self._core.extract_smartart_placement(
                *args, **kwargs
            ),
            read_smartart_reset_state=lambda *args, **kwargs: self._core.extract_smartart_reset_state(
                *args, **kwargs
            ),
        )

    # Text Effects Verification (SmartArt text / Shape text effects)

    def _verify_text_effects(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_smartart_text.verify_text_effects(
            pred_path,
            gold_path,
            spec,
            read_smartart_camera_state=lambda *args, **kwargs: self._core.extract_smartart_camera_state(
                *args, **kwargs
            ),
            read_smartart_subnode_properties=lambda *args, **kwargs: self._core.extract_smartart_subnode_properties(
                *args, **kwargs
            ),
            verify_shapes_match=lambda *args, **kwargs: self._verify_shapes_match(*args, **kwargs),
        )

    # Table Verification

    def _verify_table(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None,
        *, doc_in_path: Optional[Path] = None,
    ) -> VerificationResult:
        return families_tables_verify.verify_table(
            pred_path,
            gold_path,
            spec,
            doc_in_path=doc_in_path,
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            read_table_properties=lambda *args, **kwargs: self._core.extract_table_properties(
                *args, **kwargs
            ),
        )

    def _verify_saved_table_column_delete(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_tables_verify.verify_saved_table_column_delete(
            pred_path,
            gold_path,
        )

    _rendered_table_width_tolerances = staticmethod(
        families_tables_compare.rendered_table_width_tolerances
    )

    def _verify_table_window_distribute_align_com(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_tables_verify.verify_table_window_distribute_align_com(
            pred_path,
            gold_path,
            read_table_properties=lambda *args, **kwargs: self._core.extract_table_properties(
                *args, **kwargs
            ),
        )

    def _verify_saved_table_window_distribute_align(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_tables_verify.verify_saved_table_window_distribute_align(
            pred_path,
            gold_path,
        )

    def _verify_table_column_insert_semantics(
        self, gold_tables: List[Dict], pred_tables: List[Dict], idMso: str
    ) -> VerificationResult:
        return families_tables_compare.verify_table_column_insert_semantics(
            gold_tables,
            pred_tables,
            idMso,
        )

    def _verify_insert_then_border_semantics(
        self, gold_tables: List[Dict], pred_tables: List[Dict]
    ) -> VerificationResult:
        return families_tables_compare.verify_insert_then_border_semantics(
            gold_tables,
            pred_tables,
        )

    _verify_additive_table_style_direction_border = staticmethod(
        families_tables_compare.verify_additive_table_style_direction_border
    )

    _verify_inserted_table_properties = staticmethod(
        families_tables_compare.verify_inserted_table_properties
    )

    def _verify_table_autofit(
        self,
        gold_tables: List[Dict],
        pred_tables: List[Dict],
        pred_path: Path,
        gold_path: Path,
        idMso: str,
    ) -> VerificationResult:
        return families_tables_verify.verify_table_autofit(
            gold_tables,
            pred_tables,
            pred_path,
            gold_path,
            idMso,
        )

    # Multi-Shape Composition Verification (layout-coupled)

    def _relative_only_layout_composition(self, spec: Optional[Dict]) -> bool:
        return composition_layout.relative_only_layout_composition(
            spec,
        )

    def _verify_multi_shape_composition(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return composition_layout.verify_multi_shape_composition(
            pred_path,
            gold_path,
            spec,
            calculate_actual_position=lambda *args, **kwargs: self._calculate_actual_position(
                *args, **kwargs
            ),
            extract_shape_pair=lambda *args, **kwargs: self._extract_shape_pair(*args, **kwargs),
            match_shapes_by_name=lambda *args, **kwargs: self._match_shapes_by_name(
                *args, **kwargs
            ),
            wrap_polygon_failures=lambda *args, **kwargs: self._wrap_polygon_failures(
                *args, **kwargs
            ),
        )

    # Generic Shape Match Verification

    def _verify_shapes_match(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_verify_collection.verify_shapes_match(
            pred_path,
            gold_path,
            spec,
            get_position_tolerance=lambda: self._core.POSITION_TOLERANCE_PT,
            read_document_structure=lambda *args, **kwargs: self._core.extract_document_structure(
                *args, **kwargs
            ),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    def _verify_document_structure(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_verify_collection.verify_document_structure(
            pred_path,
            gold_path,
            read_document_structure=lambda *args, **kwargs: self._core.extract_document_structure(
                *args, **kwargs
            ),
        )

    # Helper Methods

    def _compare_shape_positions(
        self, gold_shapes: List[Dict], pred_shapes: List[Dict], tolerance: float
    ) -> Optional[VerificationResult]:
        return families_shapes_verify_collection.compare_shape_positions(
            gold_shapes,
            pred_shapes,
            tolerance,
        )
