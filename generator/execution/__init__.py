"""Execution templates for benchmark generation."""

from .setup import (
    ensure_word,
    finalize_word,
    SetupTemplate,
    SimpleTextSetup,
    MultiParagraphSetup,
    get_setup_template,
)
from .execution import (
    ExecutionTemplate,
    DirectExecute,
    get_execution_template,
)

__all__ = [
    'ensure_word',
    'finalize_word',
    'SetupTemplate',
    'SimpleTextSetup',
    'MultiParagraphSetup',
    'get_setup_template',
    'ExecutionTemplate',
    'DirectExecute',
    'get_execution_template',
]
