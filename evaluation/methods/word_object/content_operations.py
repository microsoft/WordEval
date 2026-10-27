"""
Content operations verification for WordEval.

Handles verification for content_operations behavior across all scopes:
- core_actionable (FootnoteInsert, PageBreakInsert, HeaderFooterRemove, etc.)
- core_parameterized, contextual_actionable, etc.
- specialized_actionable (Track Changes, Comments, Drawing Canvas, etc.)

Verification methods are spec-driven:
- footnote_verification: Footnote count and content matches
- endnote_verification: Endnote count and content matches
- header_footer_verification: Header/footer content matches
- section_verification: Section properties match
- page_verification: Page count and breaks match
- toc_verification: Table of Contents matches
- index_verification: Index content matches
- bibliography_verification: Bibliography content matches
- equation_verification: Equation count matches
- structure_verification: Document structure counts match

NEW - Specialized Actionable:
- revision_verification: Track changes (Accept/Reject) - checks Revisions.Count,
                         TrackRevisions, and Content.Text comparison
- comment_verification: Comment count matches gold
- comment_text_verification: Comment with specific text exists (instruction match)
- comment_reply_verification: Reply comment with specific text exists
- resolved_comment_verification: All resolved comments deleted, count check
- drawing_canvas_verification: Ink drawing canvas (msoCanvas) exists
"""

import hashlib
import io
import math
import posixpath
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from ..base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from .composition import composition_command_ids
from .core import DocumentOpenError, WordObjectVerification

from .families.charts import (
    compare as chart_compare,
    extract_saved as chart_saved,
    insert as chart_insert,
    workbook as chart_workbook,
)
from .ooxml.package import part_relationships


from .families.equations import insert as families_equations_insert
from .families.equations import insert_saved as families_equations_insert_saved
from .families.tables import insert as families_tables_insert
from .families.tables import insert_saved as families_tables_insert_saved


from .families.icons import insert as families_icons_insert
from .families.icons import insert_saved as families_icons_insert_saved
from .families.pictures import insert as families_pictures_insert
from .families.shapes import insert as families_shapes_insert
from .families.shapes import insert_geometry as families_shapes_insert_geometry
from .families.shapes import insert_saved as families_shapes_insert_saved
from .families.smartart import insert as families_smartart_insert
from .families.smartart import insert_saved as families_smartart_insert_saved


from .families.comments import compare as families_comments_compare
from .families.comments import extract_saved as families_comments_extract_saved
from .families.comments import verify as families_comments_verify
from .families.content_controls import verify as families_content_controls_verify
from .families.document_compositions import cleanup as families_document_compositions_cleanup
from .families.document_compositions import enrichment as families_document_compositions_enrichment
from .families.document_compositions import (
    extract_saved as families_document_compositions_extract_saved,
)
from .families.document_compositions import (
    preservation as families_document_compositions_preservation,
)
from .families.documents import verify as families_documents_verify
from .families.headers import verify as families_headers_verify
from .families.hyperlinks import extract_saved as families_hyperlinks_extract_saved
from .families.hyperlinks import verify as families_hyperlinks_verify
from .families.metadata import verify as families_metadata_verify
from .families.notes import compare as families_notes_compare
from .families.notes import extract_saved as families_notes_extract_saved
from .families.notes import verify as families_notes_verify
from .families.notes import verify_saved as families_notes_verify_saved
from .families.pages import extract_saved as families_pages_extract_saved
from .families.pages import numbers as families_pages_numbers
from .families.pages import verify as families_pages_verify
from .families.references import bibliography as families_references_bibliography
from .families.references import extract_saved as families_references_extract_saved
from .families.references import toc as families_references_toc
from .families.references import verify as families_references_verify
from .families.revisions import extract_saved as families_revisions_extract_saved
from .families.revisions import verify as families_revisions_verify
from .ooxml import canonical as ooxml_canonical
from .policies import content_operations as policies_content_operations
from .preservation import context as preservation_context


from . import composition_content as composition_content
from .families.formatting import operations as families_formatting_operations


class ContentOperationsVerifier:
    """
    Verifier for content_operations behavior.

    Uses WordObjectVerification core to extract document structure
    and compare them between predicted and gold documents.
    """

    def __init__(self, core: WordObjectVerification):
        self._core = core
        self._doc_in_path: Optional[Path] = None  # Set during verify()

    _error_with_details = staticmethod(policies_content_operations.error_with_details)

    def verify_saved_fallback(
        self,
        pred_path: Path,
        gold_path: Path,
        doc_in_path: Optional[Path] = None,
        spec: Optional[Dict] = None,
    ) -> Optional[VerificationResult]:
        """Return a registered saved-state verifier after a COM timeout."""
        spec = spec or {}
        self._doc_in_path = doc_in_path
        primary = spec.get("verification", {}).get("primary", {})
        method = primary.get("method", "")
        if method == "document_cleanup_composition_verification":
            return self._saved_result_after_com_failure(
                self._verify_document_cleanup_saved(pred_path, gold_path, spec)
            )
        if method == "document_enrichment_composition_verification":
            return self._saved_result_after_com_failure(
                self._verify_document_enrichment_saved(pred_path, gold_path, spec)
            )
        if method == "toc_text_level_verification":
            target_text = (
                spec.get("precondition", {}).get("select_text")
                or spec.get("precondition", {}).get("full_text")
                or "Alpha TARGET Omega."
            )
            return self._saved_result_after_com_failure(
                self._verify_saved_toc_text_level(pred_path, gold_path, target_text, spec)
            )
        if method == "resolved_comment_verification":
            expected_count = primary.get("params", {}).get("expected_count") or primary.get(
                "expected_count"
            )
            return self._saved_result_after_com_failure(
                self._verify_resolved_comments_deleted_saved(
                    pred_path, gold_path, expected_count, spec
                )
            )
        if method == "comment_verification":
            return self._saved_result_after_com_failure(
                self._verify_comments_saved(pred_path, gold_path, spec)
            )
        if method in {"comment_text_verification", "comment_reply_verification"}:
            params = primary.get("params", {})
            inferred = self._infer_comment_expectation_from_saved_gold(gold_path)
            expected_text = (
                params.get("expected_text")
                or primary.get("expected_text")
                or inferred.get("expected_text", "")
            )
            expected_author = (
                params.get("expected_author")
                or primary.get("expected_author")
                or inferred.get("expected_author")
            )
            reply = (
                method == "comment_reply_verification"
                or spec.get("variation") == "reply"
                or inferred.get("mode") == "reply"
            )
            saved_result = (
                self._verify_comment_reply_saved(
                    pred_path, gold_path, expected_text, expected_author, spec
                )
                if reply
                else self._verify_comment_text_saved(pred_path, gold_path, expected_text, spec)
            )
            return self._saved_result_after_com_failure(saved_result)
        if method == "footnote_verification":
            return self._saved_result_after_com_failure(
                self._verify_footnotes_saved(pred_path, gold_path, spec)
            )
        note_methods = {
            "endnote_verification": ("endnote", "insert"),
            "footnote_number_format_verification": ("footnote", "number_format"),
            "endnote_number_format_verification": ("endnote", "number_format"),
            "footnote_numbering_rule_verification": ("footnote", "numbering_rule"),
            "footnote_custom_symbol_verification": ("footnote", "custom_symbol"),
            "endnote_custom_symbol_verification": ("endnote", "custom_symbol"),
        }
        if method in note_methods:
            kind, operation = note_methods[method]
            return self._saved_result_after_com_failure(
                self._verify_note_saved_operation(pred_path, gold_path, kind, operation, spec)
            )
        if method == "revision_verification":
            return self._saved_result_after_com_failure(
                self._verify_revisions_saved(pred_path, gold_path, spec)
            )
        if method == "track_changes_verification":
            return self._saved_result_after_com_failure(
                self._verify_track_changes_saved(pred_path, gold_path, spec)
            )
        if method == "table_verification" and spec.get("idMso") == "TableInsertGallery":
            return self._verify_table_insert_saved(pred_path, gold_path)
        if method == "table_verification" and spec.get("idMso") == "QuickTablesInsertGallery":
            return self._saved_result_after_com_failure(
                self._verify_table_insert_saved(pred_path, gold_path)
            )
        if method == "chart_insert_verification":
            return self._saved_result_after_com_failure(
                self._verify_chart_insert_saved(pred_path, gold_path, spec)
            )
        if method == "smartart_insert_verification":
            return self._saved_result_after_com_failure(
                self._verify_smartart_insert_saved(pred_path, gold_path, spec)
            )
        if method == "icon_insert_verification":
            return self._saved_result_after_com_failure(
                self._verify_icon_insert_saved(pred_path, gold_path, spec)
            )
        return None

    _saved_result_after_com_failure = staticmethod(
        policies_content_operations.saved_result_after_com_failure
    )

    _extract_saved_enrichment_signature = staticmethod(
        families_document_compositions_extract_saved.extract_saved_enrichment_signature
    )

    def _verify_document_enrichment_saved(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_document_compositions_enrichment.verify_document_enrichment_saved(
            pred_path,
            gold_path,
            spec,
        )

    @staticmethod
    def _com_primary_with_saved_guardrail(
        com_verify,
        saved_verify,
        *,
        saved_pass_after_com_failure: bool = False,
        saved_pass_error_pattern: Optional[str] = None
    ):
        return policies_content_operations.com_primary_with_saved_guardrail(
            com_verify,
            saved_verify,
            saved_pass_after_com_failure=saved_pass_after_com_failure,
            saved_pass_error_pattern=saved_pass_error_pattern,
        )

    _canonical_saved_element = staticmethod(ooxml_canonical.canonical_saved_element)

    @staticmethod
    def _canonical_visible_element(
        element: Optional[ET.Element],
        *,
        skipped_tags: Optional[set[str]] = None,
        skipped_attribute_names: Optional[set[str]] = None
    ) -> Any:
        return ooxml_canonical.canonical_visible_element(
            element,
            skipped_tags=skipped_tags,
            skipped_attribute_names=skipped_attribute_names,
        )

    @classmethod
    def _saved_body_visible_format_signature(
        cls,
        root: ET.Element,
        *,
        exclude_hyperlink_runs: bool = False,
        hyperlink_extra_only: bool = False
    ) -> List[Dict[str, Any]]:
        return ooxml_canonical.saved_body_visible_format_signature(
            root,
            exclude_hyperlink_runs=exclude_hyperlink_runs,
            hyperlink_extra_only=hyperlink_extra_only,
        )

    _saved_picture_signatures = staticmethod(ooxml_canonical.saved_picture_signatures)

    _note_property_tokens = staticmethod(families_notes_extract_saved.note_property_tokens)

    _normalize_word_default_layout_twips = staticmethod(
        ooxml_canonical.normalize_word_default_layout_twips
    )

    _extract_saved_note_signature = staticmethod(
        families_notes_extract_saved.extract_saved_note_signature
    )

    _extract_saved_footnote_signature = staticmethod(
        families_notes_extract_saved.extract_saved_footnote_signature
    )

    @classmethod
    def _extract_saved_semantic_context(
        cls,
        docx_path: Path,
        *,
        strip_note_references: bool = False,
        strip_note_properties: Optional[str] = None,
        strip_tracked_insertions: bool = False
    ) -> Dict[str, Any]:
        return preservation_context.extract_saved_semantic_context(
            docx_path,
            strip_note_references=strip_note_references,
            strip_note_properties=strip_note_properties,
            strip_tracked_insertions=strip_tracked_insertions,
        )

    @staticmethod
    def _render_saved_revision_text(element: ET.Element, w_uri: str, *, accept: bool) -> str:
        return families_revisions_extract_saved.render_saved_revision_text(
            element,
            w_uri,
            accept=accept,
        )

    _extract_saved_revision_signature = staticmethod(
        families_revisions_extract_saved.extract_saved_revision_signature
    )

    _extract_saved_cleanup_signature = staticmethod(
        families_document_compositions_extract_saved.extract_saved_cleanup_signature
    )

    @classmethod
    def _extract_saved_textboxes(cls, docx_path: Path) -> List[Dict[str, Any]]:
        return families_shapes_insert_saved.extract_saved_textboxes(
            docx_path,
            canonical_saved_element=lambda *args, **kwargs: cls._canonical_saved_element(
                *args, **kwargs
            ),
        )

    _extract_saved_autoshape_geometries = staticmethod(
        families_shapes_insert_saved.extract_saved_autoshape_geometries
    )

    _saved_part_relationships = staticmethod(part_relationships)

    _canonical_chart_value = staticmethod(chart_workbook.canonical_value)

    _extract_saved_workbook_cells = staticmethod(chart_workbook.extract_workbook_cells)

    _chart_column_number = staticmethod(chart_workbook.column_number)

    _chart_column_name = staticmethod(chart_workbook.column_name)

    _parse_chart_range_formula = staticmethod(chart_workbook.parse_range_formula)

    _workbook_projection = staticmethod(chart_workbook.project_workbook)

    _extract_saved_workbook_grid = staticmethod(chart_workbook.extract_workbook_grid)

    _extract_saved_chart_signature = staticmethod(chart_saved.extract_insertion_signature)

    _normalize_svg_color = staticmethod(families_icons_insert_saved.normalize_svg_color)

    _parse_svg_absolute_length = staticmethod(families_icons_insert_saved.parse_svg_absolute_length)

    @classmethod
    def _extract_saved_svg_payload(cls, payload: bytes) -> Dict[str, Any]:
        return families_icons_insert_saved.extract_saved_svg_payload(
            payload,
            canonical_chart_value=lambda *args, **kwargs: cls._canonical_chart_value(
                *args, **kwargs
            ),
        )

    @classmethod
    def _extract_saved_svg_icon_signature(cls, docx_path: Path) -> Dict[str, Any]:
        return families_icons_insert_saved.extract_saved_svg_icon_signature(
            docx_path,
            canonical_chart_value=lambda *args, **kwargs: cls._canonical_chart_value(
                *args, **kwargs
            ),
            saved_part_relationships=lambda *args, **kwargs: cls._saved_part_relationships(
                *args, **kwargs
            ),
        )

    @classmethod
    def _extract_saved_smartart_signature(cls, docx_path: Path) -> Dict[str, Any]:
        return families_smartart_insert_saved.extract_saved_smartart_signature(
            docx_path,
            saved_part_relationships=lambda *args, **kwargs: cls._saved_part_relationships(
                *args, **kwargs
            ),
        )

    _extract_saved_insertion_context = staticmethod(
        preservation_context.extract_saved_insertion_context
    )

    @classmethod
    def _verify_saved_insertion_context(
        cls, pred_path: Path, gold_path: Path, *, ignore_body_visible_format: bool = False
    ) -> Optional[VerificationResult]:
        return preservation_context.verify_saved_insertion_context(
            pred_path,
            gold_path,
            ignore_body_visible_format=ignore_body_visible_format,
        )

    _extract_saved_horizontal_line_signature = staticmethod(
        families_pictures_insert.extract_saved_horizontal_line_signature
    )

    DOCINFO_COMMANDS = {
        "DocInfoFileName",
        "DocInfoAuthor",
        "DocInfoDocTitle",
    }

    _cover_page_ooxml_layout_signature = staticmethod(
        families_pages_extract_saved.cover_page_ooxml_layout_signature
    )

    def verify(
        self,
        pred_path: Path,
        gold_path: Path,
        doc_in_path: Optional[Path] = None,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        """
        Verify content operations based on spec.

        Spec-driven dispatch:
        1. Use spec's verification.primary.method if specified
        2. Fall back to command name inference
        3. Fall back to structure comparison

        Args:
            pred_path: Path to predicted document
            gold_path: Path to gold standard document
            doc_in_path: Path to input document (for delta comparisons)
            spec: Test case specification
        """
        self._doc_in_path = doc_in_path  # Store for use in verification methods
        if not spec:
            return self._verify_structure(pred_path, gold_path)

        verification = spec.get("verification", {})
        primary = verification.get("primary", {})
        method = primary.get("method", "")
        params = primary.get("params", {})

        # Dispatch based on verification method from spec
        if method == "footnote_verification":
            return self._verify_footnotes(pred_path, gold_path, spec)

        elif method == "endnote_verification":
            return self._verify_endnotes(pred_path, gold_path, spec)

        elif method in ("header_footer_verification", "header_footer_content_check"):
            # part_type can be in params or directly in primary
            part_type = params.get("part_type") or primary.get("part_type", "footer")
            # check_type/condition can be in params or directly in primary
            check_type = (
                params.get("check_type")
                or primary.get("check_type")
                or primary.get("condition", "content_match")
            )
            return self._verify_header_footer(pred_path, gold_path, part_type, check_type, spec)

        elif method == "docinfo_verification":
            return self._verify_docinfo(pred_path, gold_path, spec)

        elif method == "section_verification":
            return self._verify_sections(pred_path, gold_path)

        elif method == "page_verification":
            return self._verify_pages(pred_path, gold_path, spec)

        elif method == "toc_verification":
            return self._verify_toc(pred_path, gold_path, spec)

        elif method == "index_verification":
            return self._verify_index(pred_path, gold_path)

        elif method == "bibliography_verification":
            return self._verify_bibliography(pred_path, gold_path, spec)

        elif method == "citation_verification":
            return self._verify_citation(pred_path, gold_path, spec)

        elif method == "equation_verification":
            return self._verify_equations(pred_path, gold_path, spec)

        elif method == "text_content_verification":
            return self._verify_text_content(pred_path, gold_path, spec)

        elif method == "table_verification":
            return self._verify_tables(pred_path, gold_path, spec)

        elif method == "dropcap_verification":
            return self._verify_dropcap(pred_path, gold_path, spec)

        elif method == "wordart_verification":
            return self._verify_wordart(pred_path, gold_path, spec)

        elif method == "inline_shape_verification":
            return self._verify_inline_shapes(pred_path, gold_path)

        elif method == "content_control_verification":
            return self._verify_content_controls(pred_path, gold_path, spec)

        elif method == "structure_verification":
            return self._verify_structure(pred_path, gold_path)

        elif method == "revision_verification":
            return self._verify_revisions(pred_path, gold_path, spec)

        elif method == "comment_verification":
            return self._verify_comments(pred_path, gold_path, spec)

        elif method == "comment_text_verification":
            expected_text = params.get("expected_text") or primary.get("expected_text", "")
            inferred = self._infer_comment_expectation_from_gold(gold_path)

            if spec.get("variation") == "reply" or inferred.get("mode") == "reply":
                if not expected_text:
                    expected_text = inferred.get("expected_text", "")
                expected_author = (
                    params.get("expected_author")
                    or primary.get("expected_author")
                    or inferred.get("expected_author")
                )
                return self._verify_comment_reply(
                    pred_path, gold_path, expected_text, expected_author, spec
                )

            if not expected_text:
                expected_text = inferred.get("expected_text", "")

            return self._verify_comment_text(pred_path, gold_path, expected_text, spec)

        elif method == "comment_reply_verification":
            expected_text = params.get("expected_text") or primary.get("expected_text", "")
            expected_author = params.get("expected_author") or primary.get("expected_author")
            return self._verify_comment_reply(
                pred_path, gold_path, expected_text, expected_author, spec
            )

        elif method == "resolved_comment_verification":
            expected_count = params.get("expected_count") or primary.get("expected_count")
            return self._verify_resolved_comments_deleted(
                pred_path, gold_path, expected_count, spec
            )

        elif method == "drawing_canvas_verification":
            return self._verify_drawing_canvas(pred_path, gold_path)

        elif method == "footnote_number_format_verification":
            return self._verify_footnote_number_format(pred_path, gold_path)

        elif method == "endnote_number_format_verification":
            return self._verify_endnote_number_format(pred_path, gold_path)

        elif method == "footnote_numbering_rule_verification":
            return self._verify_footnote_numbering_rule(pred_path, gold_path)

        elif method == "footnote_custom_symbol_verification":
            return self._verify_footnote_custom_symbol(pred_path, gold_path)

        elif method == "endnote_custom_symbol_verification":
            return self._verify_endnote_custom_symbol(pred_path, gold_path)

        elif method == "hyperlink_verification":
            return self._verify_hyperlinks(pred_path, gold_path, spec)

        elif method == "paragraph_spacing_verification":
            return self._verify_paragraph_spacing(pred_path, gold_path, spec)

        elif method == "caption_verification":
            return self._verify_caption(pred_path, gold_path, spec)

        elif method == "cover_page_verification":
            return self._verify_cover_page(pred_path, gold_path)

        elif method == "page_number_field_verification":
            return self._verify_page_number_field(pred_path, gold_path, spec)

        elif method == "margin_page_numbers_verification":
            return self._verify_margin_page_numbers(pred_path, gold_path, spec)

        elif method == "textbox_gallery_verification":
            return self._verify_textbox_gallery(pred_path, gold_path)

        elif method == "textbox_insert_verification":
            return self._verify_textbox_insert(pred_path, gold_path)

        elif method == "toc_text_level_verification":
            return self._verify_toc_text_level(pred_path, gold_path, spec)

        elif method == "shape_canvas_verification":
            return self._verify_gallery_all_shapes_and_canvas(pred_path, gold_path, spec)

        elif method == "track_changes_verification":
            return self._verify_track_changes_enabled(pred_path, gold_path, spec)

        elif method == "footnote_endnote_composition_verification":
            return self._verify_footnote_endnote_composition(pred_path, gold_path)

        elif method == "header_footer_composition_verification":
            return self._verify_header_footer_composition(pred_path, gold_path)

        elif method == "document_enrichment_composition_verification":
            return self._verify_document_enrichment_composition(pred_path, gold_path, spec)

        elif method == "document_cleanup_composition_verification":
            return self._verify_document_cleanup_composition(pred_path, gold_path, spec)

        elif method == "paragraph_properties_composition_verification":
            return self._verify_paragraph_properties_composition(pred_path, gold_path)

        elif method == "chart_insert_verification":
            return self._verify_chart_insert(pred_path, gold_path, spec)

        elif method == "smartart_insert_verification":
            return self._verify_smartart_insert(pred_path, gold_path, spec)

        elif method == "icon_insert_verification":
            return self._verify_icon_insert(pred_path, gold_path, spec)

        else:
            # Fallback: infer verification from idMso command name
            return self._verify_by_command(pred_path, gold_path, spec)

    def _verify_by_command(
        self, pred_path: Path, gold_path: Path, spec: Dict
    ) -> VerificationResult:
        """
        Infer verification type from idMso command name.

        Fallback when spec doesn't specify a verification method.
        Uses command name patterns to route to appropriate verification.
        """
        idMso = spec.get("idMso", "")

        # Track Changes commands
        if idMso == "ReviewTrackChanges":
            return self._verify_track_changes_enabled(pred_path, gold_path, spec)
        elif "AcceptAllChanges" in idMso or "RejectAllChanges" in idMso:
            return self._verify_revisions(pred_path, gold_path, spec)

        # Comment commands
        elif "DeleteAllResolvedComments" in idMso:
            return self._verify_resolved_comments_deleted(pred_path, gold_path, spec=spec)
        elif "ReviewDeleteAllComments" in idMso or "ReviewDeleteComment" in idMso:
            return self._verify_comments(pred_path, gold_path, spec)
        elif "InsertNewComment" in idMso:
            # Check variation for reply
            variation = spec.get("variation", "")
            if variation == "reply":
                # Reply should be from 'user2' as per instruction
                return self._verify_comment_reply(
                    pred_path,
                    gold_path,
                    "This is a reply.",
                    expected_author="user2",
                    spec=spec,
                )
            else:
                return self._verify_comment_text(pred_path, gold_path, "Test comment.", spec)

        # Drawing commands
        elif "InsertInkDrawingCanvas" in idMso or "DrawingCanvas" in idMso:
            return self._verify_drawing_canvas(pred_path, gold_path)

        # Infer verification from command name patterns
        elif "Footnote" in idMso:
            return self._verify_footnotes(pred_path, gold_path, spec)
        elif "Endnote" in idMso:
            return self._verify_endnotes(pred_path, gold_path, spec)
        elif "PageBreak" in idMso or "BlankPage" in idMso:
            # Page-break and blank-page commands are better checked via pagination
            # than generic structure counts.
            return self._verify_pages(pred_path, gold_path, spec)
        elif "Header" in idMso or "Footer" in idMso:
            part_type = "header" if "Header" in idMso else "footer"
            # Infer check type from command: Insert->not_empty, Remove->empty
            check_type = "empty" if "Remove" in idMso else "not_empty"
            return self._verify_header_footer(pred_path, gold_path, part_type, check_type, spec)
        elif "TableOfContentsAddTextGallery" in idMso:
            return self._verify_toc_text_level(pred_path, gold_path, spec)
        elif idMso == "CoverPageInsertGallery":
            return self._verify_cover_page(pred_path, gold_path)
        elif idMso == "PageNumberFieldInsertGallery":
            return self._verify_page_number_field(pred_path, gold_path, spec)
        elif idMso == "PageNambersInMarginsInsertGallery":
            return self._verify_margin_page_numbers(pred_path, gold_path, spec)
        elif idMso == "TextBoxInsertGallery":
            return self._verify_textbox_gallery(pred_path, gold_path)
        elif idMso in ("TextBoxInsert", "TextBoxInsertVerticalWord"):
            return self._verify_textbox_insert(pred_path, gold_path)
        elif idMso == "SymbolInsertGallery":
            return self._verify_text_content(pred_path, gold_path, spec)
        elif idMso in ("QuickTablesInsertGallery", "TableInsertGallery"):
            return self._verify_tables(pred_path, gold_path, spec)
        elif idMso == "GalleryAllShapesAndCanvas":
            return self._verify_gallery_all_shapes_and_canvas(pred_path, gold_path, spec)
        elif "TableOfContents" in idMso or "TOC" in idMso:
            return self._verify_toc(pred_path, gold_path, spec)
        elif "Index" in idMso:
            return self._verify_index(pred_path, gold_path)
        elif "Equation" in idMso:
            return self._verify_equations(pred_path, gold_path, spec)
        elif "DropCap" in idMso:
            return self._verify_dropcap(pred_path, gold_path, spec)
        elif "WordArt" in idMso:
            return self._verify_wordart(pred_path, gold_path, spec)
        elif "Citation" in idMso:
            return self._verify_citation(pred_path, gold_path, spec)
        elif "Bibliography" in idMso:
            return self._verify_bibliography(pred_path, gold_path, spec)
        elif "Hyperlink" in idMso:
            return self._verify_hyperlinks(pred_path, gold_path, spec)
        elif "ParagraphSpacing" in idMso:
            return self._verify_paragraph_spacing(pred_path, gold_path, spec)
        elif idMso == "SortDialogClassic":
            return self._verify_tables(pred_path, gold_path)
        elif idMso == "CaptionInsert":
            return self._verify_caption(pred_path, gold_path, spec)
        else:
            # Fallback: compare document structure counts
            return self._verify_structure(pred_path, gold_path)

    # Body Text Verification (for deletion cases)

    def _verify_body_text_matches(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_documents_verify.verify_body_text_matches(
            pred_path,
            gold_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
        )

    def _should_verify_note_body_text(self, spec: Optional[Dict]) -> bool:
        return families_notes_compare.should_verify_note_body_text(
            spec,
        )

    # Footnote Verification

    def _normalize_note_anchor_text(self, text: Optional[str]) -> str:
        return families_notes_compare.normalize_note_anchor_text(
            text,
        )

    def _extract_note_anchor_signature(
        self, body_text: str, position: int, anchor_hint: Optional[str] = None
    ) -> Dict[str, Optional[str]]:
        return families_notes_compare.extract_note_anchor_signature(
            body_text,
            position,
            anchor_hint,
        )

    def _note_reference_positions_match(
        self,
        pred_path: Path,
        gold_path: Path,
        pred_pos: Optional[int],
        gold_pos: Optional[int],
        spec: Optional[Dict] = None,
    ) -> Dict[str, Optional[object]]:
        return families_notes_compare.note_reference_positions_match(
            pred_path,
            gold_path,
            pred_pos,
            gold_pos,
            spec,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
        )

    def _saved_note_anchor_projections_match(
        self, pred_path: Path, gold_path: Path, kind: str
    ) -> Dict[str, Any]:
        return families_notes_compare.saved_note_anchor_projections_match(
            pred_path,
            gold_path,
            kind,
        )

    _saved_note_entries = staticmethod(families_notes_compare.saved_note_entries)

    @staticmethod
    def _saved_note_property_state(
        signature: Dict[str, Any], *, ignore: Optional[str] = None, kind: Optional[str] = None
    ) -> Dict[str, Any]:
        return families_notes_compare.saved_note_property_state(
            signature,
            ignore=ignore,
            kind=kind,
        )

    _effective_saved_note_property = staticmethod(
        families_notes_compare.effective_saved_note_property
    )

    _saved_note_structure_errors = staticmethod(families_notes_compare.saved_note_structure_errors)

    _note_expected_values = staticmethod(families_notes_compare.note_expected_values)

    def _verify_saved_note_context(
        self, pred_path: Path, gold_path: Path, *, strip_note_properties: Optional[str] = None
    ) -> Optional[VerificationResult]:
        return families_notes_verify_saved.verify_saved_note_context(
            pred_path,
            gold_path,
            strip_note_properties=strip_note_properties,
            doc_in_path=self._doc_in_path,
        )

    def _verify_note_saved_operation(
        self,
        pred_path: Path,
        gold_path: Path,
        kind: str,
        operation: str,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        return families_notes_verify_saved.verify_note_saved_operation(
            pred_path,
            gold_path,
            kind,
            operation,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _compare_com_note_entries(
        self,
        pred_path: Path,
        gold_path: Path,
        pred_entries: List[Dict[str, Any]],
        gold_entries: List[Dict[str, Any]],
        kind: str,
        spec: Optional[Dict],
    ) -> Optional[VerificationResult]:
        return families_notes_compare.compare_com_note_entries(
            pred_path,
            gold_path,
            pred_entries,
            gold_entries,
            kind,
            spec,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
        )

    def _verify_note_com_operation(
        self,
        pred_path: Path,
        gold_path: Path,
        kind: str,
        operation: str,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        return families_notes_verify.verify_note_com_operation(
            pred_path,
            gold_path,
            kind,
            operation,
            spec,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    def _verify_footnotes(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_notes_verify.verify_footnotes(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    def _verify_footnotes_saved(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_notes_verify_saved.verify_footnotes_saved(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _verify_footnotes_com(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_notes_verify.verify_footnotes_com(
            pred_path,
            gold_path,
            spec,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    # Endnote Verification

    def _verify_endnotes(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_notes_verify.verify_endnotes(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    # Footnote + Endnote Composition Verification

    def _verify_footnote_endnote_composition(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_notes_verify.verify_footnote_endnote_composition(
            pred_path,
            gold_path,
            spec,
            add_composition_atomic_details=lambda *args, **kwargs: self._add_composition_atomic_details(
                *args, **kwargs
            ),
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
            summarize_sub_result=lambda *args, **kwargs: self._summarize_sub_result(
                *args, **kwargs
            ),
        )

    # Header + Footer Composition Verification

    def _verify_header_footer_composition(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_headers_verify.verify_header_footer_composition(
            pred_path,
            gold_path,
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
        )

    # Document Enrichment Composition Verification

    def _verify_document_enrichment_preservation(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_document_compositions_preservation.verify_document_enrichment_preservation(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _verify_document_enrichment_composition(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_document_compositions_enrichment.verify_document_enrichment_composition(
            pred_path,
            gold_path,
            spec,
            add_composition_atomic_details=lambda *args, **kwargs: self._add_composition_atomic_details(
                *args, **kwargs
            ),
            doc_in_path=self._doc_in_path,
            read_bibliography_info=lambda *args, **kwargs: self._core.extract_bibliography_info(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_dynamic_reference_state=lambda *args, **kwargs: self._core.extract_dynamic_reference_state(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
            summarize_sub_result=lambda *args, **kwargs: self._summarize_sub_result(
                *args, **kwargs
            ),
        )

    # Document Cleanup Composition Verification

    def _verify_document_cleanup_composition(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_document_compositions_cleanup.verify_document_cleanup_composition(
            pred_path,
            gold_path,
            spec,
            add_composition_atomic_details=lambda *args, **kwargs: self._add_composition_atomic_details(
                *args, **kwargs
            ),
            doc_in_path=self._doc_in_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
            summarize_sub_result=lambda *args, **kwargs: self._summarize_sub_result(
                *args, **kwargs
            ),
        )

    def _verify_document_cleanup_saved(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_document_compositions_cleanup.verify_document_cleanup_saved(
            pred_path,
            gold_path,
            spec,
        )

    def _verify_document_cleanup_com(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_document_compositions_cleanup.verify_document_cleanup_com(
            pred_path,
            gold_path,
            spec,
            add_composition_atomic_details=lambda *args, **kwargs: self._add_composition_atomic_details(
                *args, **kwargs
            ),
            doc_in_path=self._doc_in_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
            summarize_sub_result=lambda *args, **kwargs: self._summarize_sub_result(
                *args, **kwargs
            ),
        )

    def _summarize_sub_result(self, result: VerificationResult) -> Dict[str, Any]:
        return composition_content.summarize_sub_result(
            result,
        )

    def _add_composition_atomic_details(
        self, details: Dict[str, Any], pred_path: Path, gold_path: Path, spec: Optional[Dict]
    ) -> tuple[List[str], List[str]]:
        return composition_content.add_composition_atomic_details(
            details,
            pred_path,
            gold_path,
            spec,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
            verify_bibliography=lambda *args, **kwargs: self._verify_bibliography(*args, **kwargs),
            verify_citation=lambda *args, **kwargs: self._verify_citation(*args, **kwargs),
            verify_endnotes=lambda *args, **kwargs: self._verify_endnotes(*args, **kwargs),
            verify_footnotes=lambda *args, **kwargs: self._verify_footnotes(*args, **kwargs),
            verify_header_footer=lambda *args, **kwargs: self._verify_header_footer(
                *args, **kwargs
            ),
        )

    def _build_composition_atomic_results(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict]
    ) -> List[Dict[str, Any]]:
        return composition_content.build_composition_atomic_results(
            pred_path,
            gold_path,
            spec,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
            verify_bibliography=lambda *args, **kwargs: self._verify_bibliography(*args, **kwargs),
            verify_citation=lambda *args, **kwargs: self._verify_citation(*args, **kwargs),
            verify_endnotes=lambda *args, **kwargs: self._verify_endnotes(*args, **kwargs),
            verify_footnotes=lambda *args, **kwargs: self._verify_footnotes(*args, **kwargs),
            verify_header_footer=lambda *args, **kwargs: self._verify_header_footer(
                *args, **kwargs
            ),
        )

    def _evaluate_composition_target(
        self,
        target: Dict[str, Any],
        pred_path: Path,
        gold_path: Path,
        spec: Optional[Dict],
        cache: Dict[str, Any],
    ) -> Dict[str, Any]:
        return composition_content.evaluate_composition_target(
            target,
            pred_path,
            gold_path,
            spec,
            cache,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
            verify_bibliography=lambda *args, **kwargs: self._verify_bibliography(*args, **kwargs),
            verify_citation=lambda *args, **kwargs: self._verify_citation(*args, **kwargs),
            verify_endnotes=lambda *args, **kwargs: self._verify_endnotes(*args, **kwargs),
            verify_footnotes=lambda *args, **kwargs: self._verify_footnotes(*args, **kwargs),
            verify_header_footer=lambda *args, **kwargs: self._verify_header_footer(
                *args, **kwargs
            ),
        )

    def _verify_content_operation_target(
        self,
        prop: str,
        pred_path: Path,
        gold_path: Path,
        spec: Optional[Dict],
        target: Dict[str, Any],
    ) -> VerificationResult:
        return composition_content.verify_content_operation_target(
            prop,
            pred_path,
            gold_path,
            spec,
            target,
            verify_bibliography=lambda *args, **kwargs: self._verify_bibliography(*args, **kwargs),
            verify_citation=lambda *args, **kwargs: self._verify_citation(*args, **kwargs),
            verify_endnotes=lambda *args, **kwargs: self._verify_endnotes(*args, **kwargs),
            verify_footnotes=lambda *args, **kwargs: self._verify_footnotes(*args, **kwargs),
            verify_header_footer=lambda *args, **kwargs: self._verify_header_footer(
                *args, **kwargs
            ),
        )

    def _extract_composition_target_data(
        self, extractor: str, path: Path, cache: Dict[str, Any], side: str
    ) -> Any:
        return composition_content.extract_composition_target_data(
            extractor,
            path,
            cache,
            side,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
        )

    def _composition_target_value(
        self, data: Any, prop: str, target: Optional[Dict[str, Any]] = None
    ) -> Any:
        return composition_content.composition_target_value(
            data,
            prop,
            target,
        )

    def _header_footer_target_value(self, data: Dict[str, Any], prop: str, part: str) -> Any:
        return composition_content.header_footer_target_value(
            data,
            prop,
            part,
        )

    def _composition_values_match(self, gold_value: Any, pred_value: Any) -> bool:
        return composition_content.composition_values_match(
            gold_value,
            pred_value,
        )

    # Paragraph Properties Composition Verification

    def _verify_paragraph_properties_composition(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_formatting_operations.verify_paragraph_properties_composition(
            pred_path,
            gold_path,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
            verify_dropcap=lambda *args, **kwargs: self._verify_dropcap(*args, **kwargs),
        )

    # Footnote/Endnote Number Format Verification

    def _verify_footnote_number_format(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_notes_verify.verify_footnote_number_format(
            pred_path,
            gold_path,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    def _verify_endnote_number_format(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_notes_verify.verify_endnote_number_format(
            pred_path,
            gold_path,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    # Footnote Numbering Rule Verification

    def _verify_footnote_numbering_rule(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_notes_verify.verify_footnote_numbering_rule(
            pred_path,
            gold_path,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    # Footnote/Endnote Custom Symbol Verification

    def _verify_footnote_custom_symbol(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_notes_verify.verify_footnote_custom_symbol(
            pred_path,
            gold_path,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    def _verify_endnote_custom_symbol(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_notes_verify.verify_endnote_custom_symbol(
            pred_path,
            gold_path,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_footnote_endnote_content=lambda *args, **kwargs: self._core.extract_footnote_endnote_content(
                *args, **kwargs
            ),
        )

    # Header/Footer Verification

    def _verify_header_footer(
        self,
        pred_path: Path,
        gold_path: Path,
        part_type: str = "footer",
        check_type: str = "content_match",
        spec: Optional[Dict] = None,
        *,
        verify_body_preservation: bool = True
    ) -> VerificationResult:
        return families_headers_verify.verify_header_footer(
            pred_path,
            gold_path,
            part_type,
            check_type,
            spec,
            verify_body_preservation=verify_body_preservation,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_header_footer_content=lambda *args, **kwargs: self._core.extract_header_footer_content(
                *args, **kwargs
            ),
        )

    def _verify_docinfo(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_metadata_verify.verify_docinfo(
            pred_path,
            gold_path,
            spec,
            read_docinfo_state=lambda *args, **kwargs: self._core.extract_docinfo_state(
                *args, **kwargs
            ),
        )

    # Section Verification

    def _verify_sections(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_pages_verify.verify_sections(
            pred_path,
            gold_path,
            read_section_properties=lambda *args, **kwargs: self._core.extract_section_properties(
                *args, **kwargs
            ),
        )

    # Page Verification

    def _verify_pages(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_pages_verify.verify_pages(
            pred_path,
            gold_path,
            spec,
            get_optional_reader=lambda extractor_name: getattr(self._core, extractor_name, None),
            read_page_info=lambda *args, **kwargs: self._core.extract_page_info(*args, **kwargs),
        )

    # TOC Verification

    _extract_saved_toc_add_text_state = staticmethod(
        families_references_extract_saved.extract_saved_toc_add_text_state
    )

    def _verify_saved_toc_text_level(
        self, pred_path: Path, gold_path: Path, target_text: str, spec: Optional[Dict]
    ) -> Optional[VerificationResult]:
        return families_references_toc.verify_saved_toc_text_level(
            pred_path,
            gold_path,
            target_text,
            spec,
        )

    def _verify_toc_text_level(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_references_toc.verify_toc_text_level(
            pred_path,
            gold_path,
            spec,
            read_paragraph_outline_info_by_text=lambda *args, **kwargs: self._core.extract_paragraph_outline_info_by_text(
                *args, **kwargs
            ),
            read_toc_content=lambda *args, **kwargs: self._core.extract_toc_content(
                *args, **kwargs
            ),
        )

    def _verify_toc_text_level_com(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_references_toc.verify_toc_text_level_com(
            pred_path,
            gold_path,
            spec,
            read_paragraph_outline_info_by_text=lambda *args, **kwargs: self._core.extract_paragraph_outline_info_by_text(
                *args, **kwargs
            ),
            read_toc_content=lambda *args, **kwargs: self._core.extract_toc_content(
                *args, **kwargs
            ),
        )

    def _verify_toc(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_references_toc.verify_toc(
            pred_path,
            gold_path,
            spec,
            read_dynamic_reference_state=lambda *args, **kwargs: self._core.extract_dynamic_reference_state(
                *args, **kwargs
            ),
        )

    # Index Verification

    def _verify_index(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_references_verify.verify_index(
            pred_path,
            gold_path,
            read_dynamic_reference_state=lambda *args, **kwargs: self._core.extract_dynamic_reference_state(
                *args, **kwargs
            ),
        )

    # Bibliography Verification

    def _verify_citation(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_references_verify.verify_citation(
            pred_path,
            gold_path,
            spec,
            read_dynamic_reference_state=lambda *args, **kwargs: self._core.extract_dynamic_reference_state(
                *args, **kwargs
            ),
        )

    def _verify_bibliography(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_references_bibliography.verify_bibliography(
            pred_path,
            gold_path,
            spec,
            read_bibliography_info=lambda *args, **kwargs: self._core.extract_bibliography_info(
                *args, **kwargs
            ),
        )

    # TextBox Gallery Verification

    def _verify_textbox_gallery(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_insert.verify_textbox_gallery(
            pred_path,
            gold_path,
            canonical_saved_element=lambda *args, **kwargs: self._canonical_saved_element(
                *args, **kwargs
            ),
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            read_shape_text_styles=lambda *args, **kwargs: self._core.extract_shape_text_styles(
                *args, **kwargs
            ),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    def _verify_textbox_gallery_saved(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_insert.verify_textbox_gallery_saved(
            pred_path,
            gold_path,
            canonical_saved_element=lambda *args, **kwargs: self._canonical_saved_element(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    def _verify_textbox_gallery_com(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_insert.verify_textbox_gallery_com(
            pred_path,
            gold_path,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            read_shape_text_styles=lambda *args, **kwargs: self._core.extract_shape_text_styles(
                *args, **kwargs
            ),
        )

    # TextBox Insert Verification (TextBoxInsert, TextBoxInsertVerticalWord)

    TEXTBOX_POSITION_TOLERANCE = 2.0  # Allow 2 points tolerance for position/size

    _is_preset_shape_position = staticmethod(
        families_shapes_insert_geometry.is_preset_shape_position
    )

    _effective_shape_coordinate = staticmethod(
        families_shapes_insert_geometry.effective_shape_coordinate
    )

    _shape_usable_page_failure = staticmethod(
        families_shapes_insert_geometry.shape_usable_page_failure
    )

    def _verify_textbox_insert(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_insert.verify_textbox_insert(
            pred_path,
            gold_path,
            canonical_saved_element=lambda *args, **kwargs: self._canonical_saved_element(
                *args, **kwargs
            ),
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    def _verify_textbox_insert_saved(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_insert.verify_textbox_insert_saved(
            pred_path,
            gold_path,
            canonical_saved_element=lambda *args, **kwargs: self._canonical_saved_element(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    def _verify_textbox_insert_com(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_insert.verify_textbox_insert_com(
            pred_path,
            gold_path,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Cover Page / Page Number Verification

    def _verify_cover_page(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_pages_verify.verify_cover_page(
            pred_path,
            gold_path,
            read_content_controls=lambda *args, **kwargs: self._core.extract_content_controls(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_page_info=lambda *args, **kwargs: self._core.extract_page_info(*args, **kwargs),
        )

    def _verify_page_number_field(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_pages_numbers.verify_page_number_field(
            pred_path,
            gold_path,
            spec,
            get_extract_body_page_number_layout=lambda: getattr(
                self._core, "extract_body_page_number_layout", None
            ),
            read_document_fields=lambda *args, **kwargs: self._core.extract_document_fields(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
        )

    def _verify_margin_page_numbers(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_pages_numbers.verify_margin_page_numbers(
            pred_path,
            gold_path,
            spec,
            get_optional_reader=lambda extractor_name: getattr(self._core, extractor_name, None),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_shape_text_styles=lambda *args, **kwargs: self._core.extract_shape_text_styles(
                *args, **kwargs
            ),
        )

    # Text Content Verification

    def _verify_text_content(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_documents_verify.verify_text_content(
            pred_path,
            gold_path,
            spec,
            read_all_paragraph_character_formatting=lambda *args, **kwargs: self._core.extract_all_paragraph_character_formatting(
                *args, **kwargs
            ),
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
        )

    # Equation Verification

    _normalize_equation_text = staticmethod(families_equations_insert_saved.normalize_equation_text)

    _saved_equation_expression = staticmethod(
        families_equations_insert_saved.saved_equation_expression
    )

    @classmethod
    def _extract_saved_equation_state(
        cls,
        docx_path: Path,
        *,
        allow_equation_gallery_wrapper: bool = False,
        allow_native_equation_wrapper: bool = False
    ) -> Dict[str, Any]:
        return families_equations_insert_saved.extract_saved_equation_state(
            docx_path,
            allow_equation_gallery_wrapper=allow_equation_gallery_wrapper,
            allow_native_equation_wrapper=allow_native_equation_wrapper,
            extract_saved_insertion_context=lambda *args, **kwargs: cls._extract_saved_insertion_context(
                *args, **kwargs
            ),
            normalize_word_default_layout_twips=lambda *args, **kwargs: cls._normalize_word_default_layout_twips(
                *args, **kwargs
            ),
        )

    def _verify_equations(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_equations_insert.verify_equations(
            pred_path,
            gold_path,
            spec,
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            extract_saved_insertion_context=lambda *args, **kwargs: self._extract_saved_insertion_context(
                *args, **kwargs
            ),
            normalize_word_default_layout_twips=lambda *args, **kwargs: self._normalize_word_default_layout_twips(
                *args, **kwargs
            ),
            read_equation_info=lambda *args, **kwargs: self._core.extract_equation_info(
                *args, **kwargs
            ),
        )

    def _verify_equations_com(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_equations_insert.verify_equations_com(
            pred_path,
            gold_path,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_equation_info=lambda *args, **kwargs: self._core.extract_equation_info(
                *args, **kwargs
            ),
        )

    def _verify_equations_saved(
        self, pred_path: Path, gold_path: Path, spec: Dict[str, Any]
    ) -> VerificationResult:
        return families_equations_insert.verify_equations_saved(
            pred_path,
            gold_path,
            spec,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            extract_saved_insertion_context=lambda *args, **kwargs: self._extract_saved_insertion_context(
                *args, **kwargs
            ),
            normalize_word_default_layout_twips=lambda *args, **kwargs: self._normalize_word_default_layout_twips(
                *args, **kwargs
            ),
        )

    # Table Verification

    def _verify_tables(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_tables_insert.verify_tables(
            pred_path,
            gold_path,
            spec,
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_table_properties=lambda *args, **kwargs: self._core.extract_table_properties(
                *args, **kwargs
            ),
        )

    _extract_saved_table_insert_signature = staticmethod(
        families_tables_insert_saved.extract_saved_table_insert_signature
    )

    def _verify_table_insert_saved(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_tables_insert.verify_table_insert_saved(
            pred_path,
            gold_path,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
        )

    def _verify_tables_com(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_tables_insert.verify_tables_com(
            pred_path,
            gold_path,
            spec,
            read_table_properties=lambda *args, **kwargs: self._core.extract_table_properties(
                *args, **kwargs
            ),
        )

    # Paragraph Spacing Verification

    def _verify_paragraph_spacing(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_formatting_operations.verify_paragraph_spacing(
            pred_path,
            gold_path,
            spec,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_all_paragraphs_formatting=lambda *args, **kwargs: self._core.extract_all_paragraphs_formatting(
                *args, **kwargs
            ),
        )

    # Caption Verification

    def _verify_caption(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_references_verify.verify_caption(
            pred_path,
            gold_path,
            spec,
            read_dynamic_reference_state=lambda *args, **kwargs: self._core.extract_dynamic_reference_state(
                *args, **kwargs
            ),
        )

    # Shape Gallery Verification

    def _verify_gallery_all_shapes_and_canvas(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_insert.verify_gallery_all_shapes_and_canvas(
            pred_path,
            gold_path,
            spec,
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    def _verify_gallery_all_shapes_and_canvas_saved(
        self, pred_path: Path, gold_path: Path
    ) -> VerificationResult:
        return families_shapes_insert.verify_gallery_all_shapes_and_canvas_saved(
            pred_path,
            gold_path,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    def _verify_gallery_all_shapes_and_canvas_com(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_shapes_insert.verify_gallery_all_shapes_and_canvas_com(
            pred_path,
            gold_path,
            spec,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_shape_properties=lambda *args, **kwargs: self._core.extract_shape_properties(
                *args, **kwargs
            ),
        )

    # Drop Cap Verification

    def _verify_dropcap(
        self,
        pred_path: Path,
        gold_path: Path,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        """Use the same semantic drop-cap contract for every routing path."""
        from .content_formatting import ContentFormattingVerifier

        return ContentFormattingVerifier(self._core)._verify_dropcap(
            pred_path, gold_path, spec, doc_in_path=self._doc_in_path)

    # WordArt Verification

    def _verify_wordart(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        """
        Verify WordArt properties for content_operations insert-gallery commands.

        Reuses object_effect's WordArt-specific comparison instead of generic
        structure matching, which is too strict/noisy for gallery style variants.
        """
        from .object_effect import ObjectEffectVerifier

        return ObjectEffectVerifier(self._core)._verify_wordart(pred_path, gold_path, spec)

    # Inline Shape Verification

    def _verify_inline_shapes(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_pictures_insert.verify_inline_shapes(
            pred_path,
            gold_path,
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_inline_shapes=lambda *args, **kwargs: self._core.extract_inline_shapes(
                *args, **kwargs
            ),
        )

    def _verify_inline_shapes_com(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_pictures_insert.verify_inline_shapes_com(
            pred_path,
            gold_path,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_inline_shapes=lambda *args, **kwargs: self._core.extract_inline_shapes(
                *args, **kwargs
            ),
        )

    def _verify_inline_shapes_saved(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_pictures_insert.verify_inline_shapes_saved(
            pred_path,
            gold_path,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
        )

    # Content Control Verification

    _extract_saved_manual_toc_state = staticmethod(
        families_references_extract_saved.extract_saved_manual_toc_state
    )

    def _verify_content_controls(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_content_controls_verify.verify_content_controls(
            pred_path,
            gold_path,
            spec,
            read_content_controls=lambda *args, **kwargs: self._core.extract_content_controls(
                *args, **kwargs
            ),
            read_dynamic_reference_state=lambda *args, **kwargs: self._core.extract_dynamic_reference_state(
                *args, **kwargs
            ),
        )

    # Generic Structure Verification

    def _verify_structure(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_documents_verify.verify_structure(
            pred_path,
            gold_path,
            read_document_structure=lambda *args, **kwargs: self._core.extract_document_structure(
                *args, **kwargs
            ),
        )

    # Revision (Track Changes) Verification

    _normalized_revision_body_text = staticmethod(
        families_revisions_verify.normalized_revision_body_text
    )

    def _verify_revisions_saved(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_revisions_verify.verify_revisions_saved(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _verify_revisions_com(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_revisions_verify.verify_revisions_com(
            pred_path,
            gold_path,
            spec,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
        )

    def _verify_track_changes_saved(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_revisions_verify.verify_track_changes_saved(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _verify_track_changes_com(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_revisions_verify.verify_track_changes_com(
            pred_path,
            gold_path,
            spec,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
        )

    def _verify_revisions(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_revisions_verify.verify_revisions(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
        )

    def _verify_track_changes_enabled(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_revisions_verify.verify_track_changes_enabled(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
            read_document_text=lambda *args, **kwargs: self._core.extract_document_text(
                *args, **kwargs
            ),
            read_revision_info=lambda *args, **kwargs: self._core.extract_revision_info(
                *args, **kwargs
            ),
        )

    # Comment Verification

    _normalized_comment_text = staticmethod(families_comments_compare.normalized_comment_text)

    _normalized_comment_scope = staticmethod(families_comments_compare.normalized_comment_scope)

    _comment_semantic_entries = staticmethod(families_comments_compare.comment_semantic_entries)

    _comment_parent_signature = staticmethod(families_comments_compare.comment_parent_signature)

    _remove_comment_thread = staticmethod(families_comments_compare.remove_comment_thread)

    _find_single_added_comment = staticmethod(families_comments_compare.find_single_added_comment)

    @classmethod
    def _comment_operation_result(
        cls,
        pred_info: Dict[str, Any],
        gold_info: Dict[str, Any],
        input_info: Optional[Dict[str, Any]],
        operation: str,
        *,
        expected_text: str = "",
        expected_author: Optional[str] = None,
        expected_count: Optional[int] = None,
        match_mode: str
    ) -> VerificationResult:
        return families_comments_compare.comment_operation_result(
            pred_info,
            gold_info,
            input_info,
            operation,
            expected_text=expected_text,
            expected_author=expected_author,
            expected_count=expected_count,
            match_mode=match_mode,
        )

    def _comment_operation_com(
        self, pred_path: Path, gold_path: Path, operation: str, **kwargs
    ) -> VerificationResult:
        return families_comments_verify.comment_operation_com(
            pred_path,
            gold_path,
            operation,
            **kwargs,
            doc_in_path=self._doc_in_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
        )

    _extract_saved_comment_context = staticmethod(
        families_comments_extract_saved.extract_saved_comment_context
    )

    def _verify_saved_comment_context(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> Optional[VerificationResult]:
        return families_comments_verify.verify_saved_comment_context(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _comment_operation_saved(
        self,
        pred_path: Path,
        gold_path: Path,
        operation: str,
        *,
        spec: Optional[Dict] = None,
        **kwargs
    ) -> VerificationResult:
        return families_comments_verify.comment_operation_saved(
            pred_path,
            gold_path,
            operation,
            spec=spec,
            **kwargs,
            doc_in_path=self._doc_in_path,
        )

    def _verify_comments(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_comments_verify.verify_comments(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
        )

    def _verify_comments_saved(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_comments_verify.verify_comments_saved(
            pred_path,
            gold_path,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _verify_comment_text(
        self, pred_path: Path, gold_path: Path, expected_text: str, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_comments_verify.verify_comment_text(
            pred_path,
            gold_path,
            expected_text,
            spec,
            doc_in_path=self._doc_in_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
        )

    def _verify_comment_text_saved(
        self, pred_path: Path, gold_path: Path, expected_text: str, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_comments_verify.verify_comment_text_saved(
            pred_path,
            gold_path,
            expected_text,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _infer_comment_expectation_from_gold(self, gold_path: Path) -> Dict:
        return families_comments_verify.infer_comment_expectation_from_gold(
            gold_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
        )

    def _infer_comment_expectation_from_saved_gold(self, gold_path: Path) -> Dict[str, Any]:
        return families_comments_verify.infer_comment_expectation_from_saved_gold(
            gold_path,
        )

    def _verify_comment_reply(
        self,
        pred_path: Path,
        gold_path: Path,
        expected_text: str,
        expected_author: Optional[str] = None,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        return families_comments_verify.verify_comment_reply(
            pred_path,
            gold_path,
            expected_text,
            expected_author,
            spec,
            doc_in_path=self._doc_in_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
        )

    def _verify_comment_reply_saved(
        self,
        pred_path: Path,
        gold_path: Path,
        expected_text: str,
        expected_author: Optional[str] = None,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        return families_comments_verify.verify_comment_reply_saved(
            pred_path,
            gold_path,
            expected_text,
            expected_author,
            spec,
            doc_in_path=self._doc_in_path,
        )

    def _verify_resolved_comments_deleted(
        self,
        pred_path: Path,
        gold_path: Path,
        expected_count: Optional[int] = None,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        return families_comments_verify.verify_resolved_comments_deleted(
            pred_path,
            gold_path,
            expected_count,
            spec,
            doc_in_path=self._doc_in_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
        )

    def _verify_resolved_comments_deleted_com(
        self, pred_path: Path, gold_path: Path, expected_count: Optional[int] = None
    ) -> VerificationResult:
        return families_comments_verify.verify_resolved_comments_deleted_com(
            pred_path,
            gold_path,
            expected_count,
            doc_in_path=self._doc_in_path,
            read_comment_info=lambda *args, **kwargs: self._core.extract_comment_info(
                *args, **kwargs
            ),
        )

    def _verify_resolved_comments_deleted_saved(
        self,
        pred_path: Path,
        gold_path: Path,
        expected_count: Optional[int] = None,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        return families_comments_verify.verify_resolved_comments_deleted_saved(
            pred_path,
            gold_path,
            expected_count,
            spec,
            doc_in_path=self._doc_in_path,
        )

    _resolved_comments_deleted_result = staticmethod(
        families_comments_compare.resolved_comments_deleted_result
    )

    _extract_comment_info_ooxml = staticmethod(
        families_comments_extract_saved.extract_comment_info_ooxml
    )

    # Drawing Canvas Verification

    # Expected canvas size from instruction: 100x200 points
    EXPECTED_CANVAS_WIDTH = 432.0
    EXPECTED_CANVAS_HEIGHT = 252.0
    SIZE_TOLERANCE = 1.0  # Allow 1 point tolerance

    def _verify_drawing_canvas(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_shapes_insert.verify_drawing_canvas(
            pred_path,
            gold_path,
            read_find_shapes_by_type=lambda *args, **kwargs: self._core.find_shapes_by_type(
                *args, **kwargs
            ),
        )

    # Hyperlink Verification

    _saved_hyperlink_guardrail_signature = staticmethod(
        families_hyperlinks_extract_saved.saved_hyperlink_guardrail_signature
    )

    _verify_hyperlink_saved_guardrail = staticmethod(
        families_hyperlinks_verify.verify_hyperlink_saved_guardrail
    )

    def _verify_hyperlinks(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_hyperlinks_verify.verify_hyperlinks(
            pred_path,
            gold_path,
            spec,
            read_hyperlink_info=lambda *args, **kwargs: self._core.extract_hyperlink_info(
                *args, **kwargs
            ),
        )

    # Chart Insert Verification

    def _verify_hyperlink_remove_com(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        return families_hyperlinks_verify.verify_hyperlink_remove_com(
            pred_path,
            gold_path,
            read_hyperlink_info=lambda *args, **kwargs: self._core.extract_hyperlink_info(
                *args, **kwargs
            ),
        )

    _verify_hyperlink_remove_saved = staticmethod(
        families_hyperlinks_verify.verify_hyperlink_remove_saved
    )

    _expected_chart_insert_contract = staticmethod(chart_compare.expected_insert_contract)

    _chart_com_contract_failures = staticmethod(chart_compare.com_insert_failures)

    _chart_saved_contract_failures = staticmethod(chart_compare.saved_insert_failures)

    def _verify_chart_insert_saved(
        self, pred_path: Path, gold_path: Path, spec: Dict[str, Any]
    ) -> VerificationResult:
        return chart_insert.verify_saved(
            pred_path,
            gold_path,
            spec,
            verify_context=self._verify_saved_insertion_context,
            error_result=self._error_with_details,
        )

    _expected_smartart_insert_contract = staticmethod(
        families_smartart_insert.expected_smartart_insert_contract
    )

    _smartart_edge_map = staticmethod(families_smartart_insert.smartart_edge_map)

    _smartart_com_contract_failures = staticmethod(
        families_smartart_insert.smartart_com_contract_failures
    )

    _smartart_saved_contract_failures = staticmethod(
        families_smartart_insert.smartart_saved_contract_failures
    )

    def _verify_smartart_insert_saved(
        self, pred_path: Path, gold_path: Path, spec: Dict[str, Any]
    ) -> VerificationResult:
        return families_smartart_insert.verify_smartart_insert_saved(
            pred_path,
            gold_path,
            spec,
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            saved_part_relationships=lambda *args, **kwargs: self._saved_part_relationships(
                *args, **kwargs
            ),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    _expected_svg_icon_contract = staticmethod(families_icons_insert.expected_svg_icon_contract)

    _svg_icon_com_contract_failures = staticmethod(
        families_icons_insert.svg_icon_com_contract_failures
    )

    _svg_icon_saved_contract_failures = staticmethod(
        families_icons_insert.svg_icon_saved_contract_failures
    )

    def _verify_icon_insert_saved(
        self, pred_path: Path, gold_path: Path, spec: Dict[str, Any]
    ) -> VerificationResult:
        return families_icons_insert.verify_icon_insert_saved(
            pred_path,
            gold_path,
            spec,
            canonical_chart_value=lambda *args, **kwargs: self._canonical_chart_value(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            saved_part_relationships=lambda *args, **kwargs: self._saved_part_relationships(
                *args, **kwargs
            ),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    def _verify_chart_insert(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return chart_insert.verify_insert(
            pred_path,
            gold_path,
            spec,
            read_com=lambda path: self._core.extract_inserted_chart_state(path),
            verify_saved=self._verify_chart_insert_saved,
            combine_evidence=self._com_primary_with_saved_guardrail,
            error_result=self._error_with_details,
        )

    # SmartArt Insert Verification

    def _verify_smartart_insert(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_smartart_insert.verify_smartart_insert(
            pred_path,
            gold_path,
            spec,
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_inserted_smartart_state=lambda *args, **kwargs: self._core.extract_inserted_smartart_state(
                *args, **kwargs
            ),
            saved_part_relationships=lambda *args, **kwargs: self._saved_part_relationships(
                *args, **kwargs
            ),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )

    # Icon Insert Verification

    def _verify_icon_insert(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        return families_icons_insert.verify_icon_insert(
            pred_path,
            gold_path,
            spec,
            canonical_chart_value=lambda *args, **kwargs: self._canonical_chart_value(
                *args, **kwargs
            ),
            com_primary_with_saved_guardrail=lambda *args, **kwargs: self._com_primary_with_saved_guardrail(
                *args, **kwargs
            ),
            error_with_details=lambda *args, **kwargs: self._error_with_details(*args, **kwargs),
            read_inserted_svg_icon_state=lambda *args, **kwargs: self._core.extract_inserted_svg_icon_state(
                *args, **kwargs
            ),
            saved_part_relationships=lambda *args, **kwargs: self._saved_part_relationships(
                *args, **kwargs
            ),
            verify_saved_insertion_context=lambda *args, **kwargs: self._verify_saved_insertion_context(
                *args, **kwargs
            ),
        )
