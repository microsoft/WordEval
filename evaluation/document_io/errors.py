"""Word document-open failures and evidence ownership."""

import os
from pathlib import Path
from typing import Any, Dict, Optional


def document_path_identity(path) -> str:
    """Compare normal and extended-length spellings of the same Word path."""
    value = str(Path(path).resolve()).casefold()
    if value.startswith("\\\\?\\unc\\"):
        return "\\\\" + value[8:]
    if value.startswith("\\\\?\\"):
        return value[4:]
    return value


class NormalOpenError(RuntimeError):
    """One failed opening attempt, before any document/environment attribution."""

    def __init__(self, path, cause):
        self.docx_path = os.path.abspath(path)
        self.cause = cause
        super().__init__(str(cause))


class DocumentOpenError(RuntimeError):
    """Typed Word document-open failure with path and repair ownership."""

    def __init__(
        self,
        docx_path: str,
        message: str,
        *,
        normal_error: Optional[BaseException] = None,
        repair_error: Optional[BaseException] = None,
        repair_succeeded: bool = False,
    ):
        super().__init__(message)
        self.docx_path = os.path.abspath(docx_path)
        self.normal_error = normal_error
        self.repair_error = repair_error
        self.repair_succeeded = repair_succeeded

    def to_details(self) -> Dict[str, Any]:
        return {
            "path": self.docx_path,
            "normal_opened": False,
            "repair_opened": self.repair_succeeded,
            "state": ("repair_required" if self.repair_succeeded else "unopenable"),
            "normal_error": (str(self.normal_error) if self.normal_error is not None else None),
            "repair_error": (str(self.repair_error) if self.repair_error is not None else None),
        }
