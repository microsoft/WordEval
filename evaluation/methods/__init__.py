"""
Verification methods for WordEval.

Two verification methods are available:
- word_object: Read properties via Word object model (Windows only)
- word_revision: Use Word's CompareDocuments API (Windows only)

Usage:
    from evaluation.methods import get_method, list_methods

    method = get_method('word_object')
    result = method.verify(pred_path, gold_path)
"""

from importlib import import_module, util
from .base import VerificationMethod, VerificationResult

# Registry of available methods
_METHODS = {}

# Root cause when method registration failed (e.g. pywin32 missing or broken).
_METHOD_MODULES = {'word_object': '.word_object', 'word_revision': '.word_revision'}
_METHOD_IMPORT_ERRORS = {}


def register_method(method_class):
    """Register a verification method class."""
    instance = method_class()
    _METHODS[instance.name] = instance
    return method_class


def get_method(name: str) -> VerificationMethod:
    """
    Get a verification method by name.

    Args:
        name: Method name ('word_object', 'word_revision')

    Returns:
        VerificationMethod instance

    Raises:
        ValueError: If method name is not recognized
    """
    if name not in _METHODS and name in _METHOD_MODULES:
        try:
            import_module(_METHOD_MODULES[name], __name__)
        except ImportError as exc:
            _METHOD_IMPORT_ERRORS[name] = f'{type(exc).__name__}: {exc}'
    if name not in _METHODS:
        available = ', '.join(list_methods())
        message = f"Unknown verification method: {name}. Available: {available}"
        if name in _METHOD_IMPORT_ERRORS:
            message += f" (method registration failed: {_METHOD_IMPORT_ERRORS[name]})"
        raise ValueError(message)
    return _METHODS[name]


def list_methods() -> list:
    """List all available verification method names."""
    return list(dict.fromkeys([
        *[name for name, module in _METHOD_MODULES.items()
          if util.find_spec(module, __name__) is not None],
        *_METHODS,
    ]))


# Load only the requested method. Optional experiment imports cannot affect
# the core benchmark or the method-independent document-access worker.


__all__ = [
    'VerificationMethod',
    'VerificationResult',
    'get_method',
    'list_methods',
    'register_method',
]
