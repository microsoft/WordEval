"""Base classes for verification methods."""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Any, Optional


@dataclass
class VerificationResult:
    """Result from a single verification method.

    Attributes:
        passed: Whether verification passed
        message: Human-readable result message
        details: Additional details (method-specific)
        error: Error message if verification couldn't run (None = ran successfully)
    """
    passed: bool
    message: str
    details: Dict[str, Any] = field(default_factory=dict)
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            'passed': self.passed,
            'message': self.message,
            'details': self.details,
            'error': self.error,
        }

    @classmethod
    def from_legacy(cls, legacy_result: Dict) -> 'VerificationResult':
        """Convert from legacy verification result format.

        Legacy format: {'passed': bool, 'message': str, 'details': dict}
        """
        return cls(
            passed=legacy_result.get('passed', False),
            message=legacy_result.get('message', ''),
            details=legacy_result.get('details', {}),
            error=legacy_result.get('error'),
        )

    @classmethod
    def error_result(cls, error_message: str, **details: Any) -> 'VerificationResult':
        """Create a result for when verification couldn't run."""
        return cls(
            passed=False,
            message=f'ERROR: {error_message}',
            details=details,
            error=error_message,
        )


def tag_verification_source(
    result: VerificationResult,
    source: str,
    **extra: Any,
) -> VerificationResult:
    """Attach backend provenance without changing the underlying verdict."""
    result.details = dict(result.details or {})
    result.details['verification_source'] = source
    result.details.update(extra)
    return result


def combine_com_primary_with_ooxml_guardrail(
    com_result: VerificationResult,
    ooxml_result: VerificationResult,
) -> VerificationResult:
    """Require a COM semantic pass and a saved-state OOXML guardrail pass.

    Word Object Model remains the primary semantic evaluator.  The OOXML
    result is retained as a secondary persistence/renderability invariant and
    may fail a document only when that saved-state property is part of the
    requested outcome.
    """
    if not com_result.passed:
        return tag_verification_source(com_result, 'word_com')
    if not ooxml_result.passed:
        return tag_verification_source(
            ooxml_result,
            'word_com+saved_ooxml_guardrail',
            word_com_primary=com_result.to_dict(),
        )
    return VerificationResult(
        passed=True,
        message=com_result.message,
        details={
            **dict(com_result.details or {}),
            'verification_source': 'word_com+saved_ooxml_guardrail',
            'saved_ooxml_guardrail': ooxml_result.to_dict(),
        },
    )


class VerificationMethod(ABC):
    """Base class for all verification methods.

    Verification methods compare predicted documents against gold standards
    using different techniques:

    - word_object: Read properties via Word object model (Windows only)
    - word_revision: Use Word's CompareDocuments API (Windows only)
    """

    @property
    @abstractmethod
    def name(self) -> str:
        """Method identifier: 'word_object', 'word_revision'"""
        pass

    @property
    def requires_word(self) -> bool:
        """Whether this method needs Microsoft Word installed."""
        return False

    @abstractmethod
    def verify(
        self,
        pred_path: Path,
        gold_path: Path,
        doc_in_path: Optional[Path] = None,
        spec: Optional[Dict] = None
    ) -> VerificationResult:
        """
        Compare predicted document against gold standard.

        Args:
            pred_path: Path to predicted document (agent output)
            gold_path: Path to gold standard document
            doc_in_path: Path to input document (optional, for some verification types)
            spec: Verification specification from test case (optional)

        Returns:
            VerificationResult with pass/fail status and details
        """
        pass

    def supports_spec(self, spec: Optional[Dict]) -> bool:
        """
        Check if this method can handle the given verification spec.

        Override in subclasses to indicate unsupported verification types.

        Args:
            spec: Verification specification from test case

        Returns:
            True if this method supports the spec, False otherwise
        """
        return True

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(name={self.name!r}, requires_word={self.requires_word})"
