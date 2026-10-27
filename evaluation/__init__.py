"""
Evaluation module for WordEval.

This module provides:
- Verification methods (word_object, word_revision)
- Evaluator class for running evaluations with multiple methods
- Unified Report format with incremental updates

Usage:
    from evaluation import Evaluator, Report
    from evaluation.methods import get_method, list_methods

CLI:
    python -m evaluation.cli -t test_set -r results/ --methods word_object word_revision
"""

from .evaluator import (
    Evaluator,
    Report,
    MethodSummary,
    load_test_cases,
)
from .methods import get_method, list_methods

__all__ = [
    'Evaluator',
    'Report',
    'MethodSummary',
    'load_test_cases',
    'get_method',
    'list_methods',
]
