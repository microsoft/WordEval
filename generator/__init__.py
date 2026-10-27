"""
Benchmark Generator Module

Generates test cases for Word commands based on BEHAVIOR × SCOPE classification.
"""

from .config import (
    PROJECT_ROOT,
    COMMAND_SETS_DIR,
    BENCHMARK_OUTPUT_DIR,
    NAMESPACES,
    TOGGLE_VARIATIONS,
    ONE_WAY_VARIATIONS,
)

__all__ = [
    'PROJECT_ROOT',
    'COMMAND_SETS_DIR',
    'BENCHMARK_OUTPUT_DIR',
    'NAMESPACES',
    'TOGGLE_VARIATIONS',
    'ONE_WAY_VARIATIONS',
]
