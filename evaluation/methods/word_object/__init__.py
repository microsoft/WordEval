"""
Word object verification package.

This package provides Word object model-based verification for WordEval.

Structure:
- core.py: WordObjectVerification - COM extraction layer
- content_formatting.py: ContentFormattingVerifier - text formatting verification
- object_effect.py: ObjectEffectVerifier - shape/object verification
- content_operations.py: ContentOperationsVerifier - document structure verification
- method.py: WordObjectMethod - entry point and behavior dispatch

Usage:
    from evaluation.methods import get_method
    method = get_method('word_object')
    result = method.verify(pred_path, gold_path, spec=spec)
"""

from .method import WordObjectMethod as _WordObjectMethodBase
from .core import WordObjectVerification
from .content_formatting import ContentFormattingVerifier
from .object_effect import ObjectEffectVerifier
from .content_operations import ContentOperationsVerifier

from .. import register_method


@register_method
class WordObjectMethod(_WordObjectMethodBase):
    """Word object verification method (registered with evaluation framework)."""

    pass


__all__ = [
    "WordObjectMethod",
    "WordObjectVerification",
    "ContentFormattingVerifier",
    "ObjectEffectVerifier",
    "ContentOperationsVerifier",
]
