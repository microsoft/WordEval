"""
Word object verification method - entry point.

This module provides the WordObjectMethod class that integrates with
the evaluation framework and dispatches to behavior-specific verifiers.

Dispatch flow:
1. Entry: WordObjectMethod.verify(pred_path, gold_path, spec)
2. Level 1: Dispatch by spec['behavior'] to behavior module
3. Level 2: Behavior module dispatches by spec['verification']['primary']['method']
"""

from pathlib import Path
from typing import Dict, Optional

from ..base import VerificationMethod, VerificationResult
from ...document_io.open_gate import normal_open_gate
from ...document_io.policy import document_open_failure_result
from .core import DocumentOpenError, WordObjectVerification
from .content_formatting import ContentFormattingVerifier
from .object_effect import ObjectEffectVerifier
from .content_operations import ContentOperationsVerifier


class WordObjectMethod(VerificationMethod):
    """
    Word object verification method.

    Uses Word's object model (via COM API) to read document properties
    and compare them between predicted and gold documents.

    Verification strategy is determined by behavior type from spec:
    - content_formatting: Compare text formatting properties on target text
    - content_operations: Compare document structure (footnotes, pages, shapes, etc.)
    - object_effect: Compare shape properties (positions, grouping, z-order, etc.)

    Each behavior module uses spec-driven dispatch to select the appropriate
    verification method based on spec['verification']['primary']['method'].
    """

    def __init__(self, max_retries: int = 3, retry_delay: float = 1.0):
        """
        Initialize word object verification method.

        Args:
            max_retries: Maximum retry attempts for COM operations
            retry_delay: Delay between retries in seconds
        """
        # Core extraction layer (shared by all verifiers)
        self._core = WordObjectVerification(max_retries=max_retries, retry_delay=retry_delay)

        # Behavior-specific verifiers
        self._content_formatting = ContentFormattingVerifier(self._core)
        self._object_effect = ObjectEffectVerifier(self._core)
        self._content_operations = ContentOperationsVerifier(self._core)

    @property
    def name(self) -> str:
        return "word_object"

    @property
    def requires_word(self) -> bool:
        return True

    def verify_normal_open_gate(self, pred_path: Path, gold_path: Path) -> VerificationResult:
        """Compatibility adapter for the shared document validity gate."""
        return VerificationResult.from_legacy(normal_open_gate(
            pred_path, gold_path, probe=WordObjectVerification.probe_clean_openability,
        ))

    def verify(
        self,
        pred_path: Path,
        gold_path: Path,
        doc_in_path: Optional[Path] = None,
        spec: Optional[Dict] = None,
    ) -> VerificationResult:
        """
        Compare predicted document against gold standard using Word object model.

        Dispatches to behavior-specific verifier based on spec['behavior']:
        - content_formatting → ContentFormattingVerifier
        - content_operations → ContentOperationsVerifier
        - object_effect → ObjectEffectVerifier

        Each behavior verifier then dispatches based on
        spec['verification']['primary']['method'].

        Args:
            pred_path: Path to predicted document (agent output)
            gold_path: Path to gold standard document
            doc_in_path: Path to input document (optional)
            spec: Test case specification with behavior and verification info

        Returns:
            VerificationResult with pass/fail status and details
        """
        try:
            # Get behavior from spec
            behavior = spec.get("behavior") if spec else None

            # Dispatch to behavior-specific verifier
            if behavior == "content_formatting":
                return self._content_formatting.verify(pred_path, gold_path, spec, doc_in_path=doc_in_path)

            elif behavior == "object_effect":
                return self._object_effect.verify(pred_path, gold_path, spec, doc_in_path=doc_in_path)

            elif behavior == "content_operations":
                return self._content_operations.verify(pred_path, gold_path, doc_in_path, spec)

            elif behavior == "selection_effect":
                return self._content_operations.verify(pred_path, gold_path, doc_in_path, spec)

            else:
                # Fallback: try to infer behavior from spec or compare structure
                return self._verify_fallback(pred_path, gold_path, spec)

        except DocumentOpenError as e:
            return VerificationResult.from_legacy(document_open_failure_result(
                e, pred_path, gold_path, doc_in_path,
            ))
        except Exception as e:
            return VerificationResult.error_result(str(e))

    def _verify_fallback(
        self, pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
    ) -> VerificationResult:
        """
        Fallback verification when behavior is not specified.

        Attempts to infer verification type from spec or falls back
        to document structure comparison.
        """
        if spec:
            # Try to infer from idMso command name
            idMso = spec.get("idMso", "")

            # Check for formatting commands
            formatting_keywords = [
                "Bold",
                "Italic",
                "Font",
                "Align",
                "Indent",
                "Strikethrough",
                "Subscript",
                "Superscript",
                "Shading",
                "Border",
                "Style",
            ]
            if any(kw in idMso for kw in formatting_keywords):
                return self._content_formatting.verify(pred_path, gold_path, spec)

            # Check for object effect commands
            object_keywords = [
                "Object",
                "Shape",
                "Align",
                "Distribute",
                "Group",
                "Rotate",
                "Flip",
                "Position",
                "Wrap",
            ]
            if any(kw in idMso for kw in object_keywords):
                return self._object_effect.verify(pred_path, gold_path, spec)

            # Check for content operation commands
            operation_keywords = [
                "Insert",
                "Remove",
                "Delete",
                "Header",
                "Footer",
                "Footnote",
                "Endnote",
                "TOC",
                "Index",
                "Break",
                "Section",
                "Page",
                "Bibliography",
                "Equation",
                "Hyperlink",
            ]
            if any(kw in idMso for kw in operation_keywords):
                # verify() signature is (pred, gold, doc_in_path=None, spec=None);
                # passing spec positionally lands it in doc_in_path and silently
                # degrades verification to a bare structure comparison.
                return self._content_operations.verify(pred_path, gold_path, spec=spec)

        # Ultimate fallback: compare document structure
        return self._content_operations._verify_structure(pred_path, gold_path)

    def verify_saved_fallback(
        self,
        pred_path: Path,
        gold_path: Path,
        doc_in_path: Optional[Path] = None,
        spec: Optional[Dict] = None,
    ) -> Optional[VerificationResult]:
        """Use an explicitly registered OOXML verifier after a COM timeout."""
        behavior = (spec or {}).get("behavior")
        if behavior == "content_formatting":
            return self._content_formatting.verify_saved_fallback(pred_path, gold_path, spec)
        if behavior == "object_effect":
            return self._object_effect.verify_saved_fallback(pred_path, gold_path, spec)
        if behavior in {"content_operations", "selection_effect"}:
            return self._content_operations.verify_saved_fallback(
                pred_path, gold_path, doc_in_path, spec
            )
        return None

    def cleanup(self):
        """Clean up Word application."""
        self._core._cleanup_word()

    def restart_word(self):
        """Restart Word application."""
        return self._core.restart_word()

    def is_word_responsive(self) -> bool:
        """Check if Word is still responsive."""
        return self._core.is_word_responsive()

    def __del__(self):
        """Cleanup on destruction."""
        self.cleanup()
