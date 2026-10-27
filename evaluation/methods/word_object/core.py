"""
Core Word object extraction layer.

WordObjectVerification preserves the extraction API, delegates feature readers
to families, and inherits process/document ownership from WordSession.
Behavior entrypoints route verification to the same feature families.
"""

import hashlib
import os
import re
import shutil
import struct
import tempfile
import time
import zipfile
import zlib
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any

from .families.charts import extract_com as chart_com
from .errors import DocumentOpenError
from .session import WordSession
from .ooxml.text import _xml_text_by_local_name
from .families.shapes.colors import _ole_rgb_channels, _has_visible_gradient_progression


from .families.equations import extract_com as families_equations_extract_com
from .families.tables import extract_com as families_tables_extract_com


from .families.icons import extract_com as families_icons_extract_com
from .families.icons import extract_saved as families_icons_extract_saved
from .families.pictures import extract_com as families_pictures_extract_com
from .families.shapes import extract_com as families_shapes_extract_com
from .families.smartart import extract_com as families_smartart_extract_com
from .families.smartart import refresh_projection as families_smartart_refresh_projection
from .families.smartart import saved_projection as families_smartart_saved_projection


from .families.comments import extract_com as families_comments_extract_com
from .families.content_controls import extract_com as families_content_controls_extract_com
from .families.documents import extract_com as families_documents_extract_com
from .families.headers import extract_com as families_headers_extract_com
from .families.hyperlinks import extract_com as families_hyperlinks_extract_com
from .families.notes import extract_com as families_notes_extract_com
from .families.pages import extract_com as families_pages_extract_com
from .families.references import collections as families_references_collections
from .families.revisions import extract_com as families_revisions_extract_com


from . import com_collections as com_collections
from .families.charts import rendered_image as families_charts_rendered_image
from .families.formatting import extract_com as families_formatting_extract_com
from .families.formatting import paragraph_com as families_formatting_paragraph_com
from .families.metadata import extract_com as families_metadata_extract_com
from .families.pages import rendered_layout as families_pages_rendered_layout
from .families.references import extract_com as families_references_extract_com
from .families.references import fields as families_references_fields


class WordObjectVerification(WordSession):
    """
    Word object verification core - extraction layer.

    Uses Word's object model (via COM API) to extract document properties.
    This class focuses on extraction; verification logic is in behavior modules.
    """

    # Tolerance for floating point comparisons (in points)
    POSITION_TOLERANCE_PT = 1.0

    # DOCX Package Helpers

    def extract_svg_icon_signatures(self, docx_path: str) -> List[Dict[str, Any]]:
        return families_icons_extract_saved.extract_svg_icon_signatures(
            docx_path,
        )

    def extract_svg_icon_properties(self, docx_path: str) -> Dict[str, Any]:
        return families_icons_extract_com.extract_svg_icon_properties(
            docx_path,
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    _strict_com_sequence = staticmethod(com_collections.strict_com_sequence)

    _png_nonwhite_ratio = staticmethod(families_charts_rendered_image.png_nonwhite_ratio)

    def extract_inserted_chart_state(self, docx_path: str) -> Dict:
        return chart_com.extract_inserted_state(
            docx_path,
            open_document=self._open_document,
            close_document=self._close_document,
            strict_sequence=self._strict_com_sequence,
            png_nonwhite_ratio=self._png_nonwhite_ratio,
        )

    def extract_inserted_smartart_state(self, docx_path: str) -> Dict[str, Any]:
        return families_smartart_extract_com.extract_inserted_smartart_state(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_inserted_svg_icon_state(self, docx_path: str) -> Dict[str, Any]:
        return families_icons_extract_com.extract_inserted_svg_icon_state(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def _classify_svg_icon_shape(self, svg_text: str) -> str:
        return families_icons_extract_saved.classify_svg_icon_shape(
            svg_text,
        )

    # Word Application Management

    # Document Operations

    # Dummy write password for writable opens: a write-reserved document
    # (w:writeProtection) prompts for its write password even with
    # DisplayAlerts=0, and the invisible dialog hangs until the case timeout.
    # Supplying a wrong password turns the prompt into an immediate COM error
    # (fail-closed). Word ignores the argument for unprotected documents.

    # Text Content Extraction

    def extract_document_text(self, docx_path: str) -> str:
        return families_documents_extract_com.extract_document_text(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_citation_fields(self, docx_path: str) -> List[Dict]:
        return families_references_fields.extract_citation_fields(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_paragraph_outline_info_by_text(
        self, docx_path: str, target_text: str
    ) -> Optional[Dict]:
        return families_formatting_paragraph_com.extract_paragraph_outline_info_by_text(
            docx_path,
            target_text,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_text_formatting(self, docx_path: str, target_text: str) -> Optional[Dict]:
        return families_formatting_extract_com.extract_text_formatting(
            docx_path,
            target_text,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            get_word=lambda *args, **kwargs: self._get_word(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_target_glyph_state(self, docx_path: str, target_text: str) -> Dict[str, Any]:
        return families_formatting_extract_com.extract_target_glyph_state(
            docx_path,
            target_text,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def _extract_fields_from_document(self, doc) -> List[Dict]:
        return families_references_fields.extract_fields_from_document(
            doc,
        )

    def extract_document_fields(self, docx_path: str, update_fields: bool = False) -> List[Dict]:
        return families_references_fields.extract_document_fields(
            docx_path,
            update_fields,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    _word_style_name = staticmethod(families_references_fields.word_style_name)

    _word_scalar = staticmethod(families_references_fields.word_scalar)

    def _update_all_word_fields(self, doc) -> List[str]:
        return families_references_fields.update_all_word_fields(
            doc,
        )

    _word_range_identity = staticmethod(families_references_fields.word_range_identity)

    def _iter_word_story_ranges(self, doc):
        return families_references_fields.iter_word_story_ranges(
            doc,
        )

    def extract_docinfo_state(self, docx_path: str, update_fields: bool = False) -> Dict[str, Any]:
        return families_metadata_extract_com.extract_docinfo_state(
            docx_path,
            update_fields,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_dynamic_reference_state(
        self, docx_path: str, update_fields: bool = False
    ) -> Dict[str, Any]:
        return families_references_extract_com.extract_dynamic_reference_state(
            docx_path,
            update_fields,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_body_page_number_layout(self, docx_path: str) -> Dict[str, Any]:
        return families_pages_rendered_layout.extract_body_page_number_layout(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_bibliography_info(self, docx_path: str) -> Dict:
        return families_references_extract_com.extract_bibliography_info(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def _extract_paragraph_borders(self, para) -> Dict[str, Optional[int]]:
        return families_formatting_paragraph_com.extract_paragraph_borders(
            para,
        )

    def _extract_paragraph_dropcap(self, para) -> Dict[str, Optional[int]]:
        return families_formatting_paragraph_com.extract_paragraph_dropcap(
            para,
        )

    def extract_all_paragraphs_formatting(self, docx_path: str) -> Optional[List[Dict]]:
        return families_formatting_paragraph_com.extract_all_paragraphs_formatting(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_paragraph_formatting(self, docx_path: str, paragraph_index: int) -> Optional[Dict]:
        return families_formatting_paragraph_com.extract_paragraph_formatting(
            docx_path,
            paragraph_index,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Drop Cap Extraction

    def extract_dropcap_properties(self, docx_path: str) -> Dict:
        return families_formatting_paragraph_com.extract_dropcap_properties(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Shape Extraction

    def extract_shape_properties(self, docx_path: str) -> List[Dict]:
        return families_shapes_extract_com.extract_shape_properties(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Header/Footer Extraction

    def extract_header_footer_content(self, docx_path: str) -> Dict:
        return families_headers_extract_com.extract_header_footer_content(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Body Paragraph Style Extraction

    def extract_all_paragraph_font_properties(self, docx_path: str) -> List[Dict]:
        return families_formatting_extract_com.extract_all_paragraph_font_properties(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_all_paragraph_character_formatting(self, docx_path: str) -> List[Dict[str, Any]]:
        return families_formatting_extract_com.extract_all_paragraph_character_formatting(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_body_paragraph_styles(self, docx_path: str) -> List[str]:
        return families_formatting_extract_com.extract_body_paragraph_styles(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Shape Text Style Extraction

    def extract_shape_text_styles(
        self, docx_path: str, include_header_footer: bool = False, include_group_items: bool = False
    ) -> List[Dict]:
        return families_shapes_extract_com.extract_shape_text_styles(
            docx_path,
            include_header_footer,
            include_group_items,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # List Extraction

    def extract_list_properties(self, docx_path: str, target_text: str = None) -> Optional[Dict]:
        return families_formatting_extract_com.extract_list_properties(
            docx_path,
            target_text,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Footnote/Endnote Extraction

    def extract_footnote_endnote_content(self, docx_path: str) -> Dict:
        return families_notes_extract_com.extract_footnote_endnote_content(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Section/Page Extraction

    def extract_section_properties(self, docx_path: str) -> List[Dict]:
        return families_pages_extract_com.extract_section_properties(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_page_info(self, docx_path: str) -> Dict:
        return families_pages_extract_com.extract_page_info(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Equation Extraction

    def extract_equation_info(self, docx_path: str) -> Dict:
        return families_equations_extract_com.extract_equation_info(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Inline Shape Extraction

    def extract_inline_shapes(self, docx_path: str) -> Dict:
        return families_pictures_extract_com.extract_inline_shapes(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_picture_crop_properties(self, docx_path: str, shape_index: int = 1) -> Dict:
        return families_pictures_extract_com.extract_picture_crop_properties(
            docx_path,
            shape_index,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_picture_format_properties(self, docx_path: str, shape_index: int = 1) -> Dict:
        return families_pictures_extract_com.extract_picture_format_properties(
            docx_path,
            shape_index,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Content Control Extraction

    def extract_content_controls(self, docx_path: str) -> Dict:
        return families_content_controls_extract_com.extract_content_controls(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Index Extraction

    def extract_index_info(self, docx_path: str) -> Dict:
        return families_references_collections.extract_index_info(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Table Extraction

    def extract_table_properties(self, docx_path: str, *, include_cell_borders: bool = False) -> List[Dict]:
        return families_tables_extract_com.extract_table_properties(
            docx_path,
            include_cell_borders=include_cell_borders,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # TOC Extraction

    def extract_toc_content(self, docx_path: str, update_fields: bool = False) -> Dict:
        return families_references_collections.extract_toc_content(
            docx_path,
            update_fields,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Document Structure Extraction

    def extract_document_structure(self, docx_path: str) -> Dict:
        return families_documents_extract_com.extract_document_structure(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Revision (Track Changes) Extraction

    def extract_revision_info(self, docx_path: str) -> Dict:
        return families_revisions_extract_com.extract_revision_info(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Comment Extraction (Extended)

    def extract_comment_info(self, docx_path: str) -> Dict:
        return families_comments_extract_com.extract_comment_info(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Hyperlink Extraction

    def extract_hyperlink_info(self, docx_path: str) -> Dict:
        return families_hyperlinks_extract_com.extract_hyperlink_info(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Shape Type Search

    def find_shapes_by_type(self, docx_path: str, shape_type: int) -> Dict:
        return families_shapes_extract_com.find_shapes_by_type(
            docx_path,
            shape_type,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # SmartArt Extraction (inline + floating)

    def extract_smartart_properties(self, docx_path: str) -> Dict:
        return families_smartart_extract_com.extract_smartart_properties(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_smartart_subnode_properties(self, docx_path: str) -> Dict:
        return families_smartart_extract_com.extract_smartart_subnode_properties(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_smartart_placement(self, docx_path: str, heading_text: str) -> Dict:
        return families_smartart_extract_com.extract_smartart_placement(
            docx_path,
            heading_text,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_shape_placement(
        self, docx_path: str, heading_text: str, auto_shape_type: Optional[int] = None
    ) -> Dict[str, Any]:
        return families_shapes_extract_com.extract_shape_placement(
            docx_path,
            heading_text,
            auto_shape_type,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_smartart_camera_state(self, docx_path: str) -> Dict[str, Any]:
        return families_smartart_saved_projection.extract_smartart_camera_state(
            docx_path,
        )

    def extract_smartart_camera_presets(self, docx_path: str) -> List[str]:
        return families_smartart_saved_projection.extract_smartart_camera_presets(
            docx_path,
        )

    def extract_smartart_drawing_info(self, docx_path: str) -> Dict:
        return families_smartart_saved_projection.extract_smartart_drawing_info(
            docx_path,
        )

    def extract_smartart_gallery_identities(self, docx_path: str) -> Dict[str, Any]:
        return families_smartart_saved_projection.extract_smartart_gallery_identities(
            docx_path,
        )

    def extract_smartart_visible_projection(self, docx_path: str) -> Dict[str, Any]:
        return families_smartart_refresh_projection.extract_smartart_visible_projection(
            docx_path,
            create_session=WordObjectVerification,
        )

    def extract_smartart_node_scale_info(self, docx_path: str) -> Dict:
        return families_smartart_saved_projection.extract_smartart_node_scale_info(
            docx_path,
        )

    def extract_smartart_node_details(self, docx_path: str) -> Dict:
        return families_smartart_extract_com.extract_smartart_node_details(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    def extract_smartart_reset_state(self, docx_path: str) -> Dict[str, Any]:
        return families_smartart_extract_com.extract_smartart_reset_state(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Chart Extraction

    def extract_chart_properties(self, docx_path: str) -> Dict:
        return chart_com.extract_properties(
            docx_path,
            open_document=self._open_document,
            close_document=self._close_document,
        )

    # Picture Effect Extraction

    def extract_picture_effect_properties(
        self, docx_path: str, shape_index: int = 1, allow_floating: bool = False
    ) -> Dict:
        return families_pictures_extract_com.extract_picture_effect_properties(
            docx_path,
            shape_index,
            allow_floating,
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Equation Structure Extraction

    def _extract_omath_func_info(self, func, index: int, depth: int = 0) -> Dict:
        return families_equations_extract_com.extract_omath_func_info(
            func,
            index,
            depth,
        )

    def extract_equation_structure(self, docx_path: str) -> Dict:
        return families_equations_extract_com.extract_equation_structure(
            docx_path,
            close_document=lambda *args, **kwargs: self._close_document(*args, **kwargs),
            open_document=lambda *args, **kwargs: self._open_document(*args, **kwargs),
        )

    # Cleanup
