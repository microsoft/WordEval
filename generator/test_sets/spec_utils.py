"""Utility for writing spec.json files with skip-if-unchanged logic.

When regenerating a test set, most spec.json files don't actually change —
only `generated_at` gets a new timestamp. This causes hundreds of noisy
git diffs. This module provides `write_spec_if_changed()` which compares
the new spec against the existing file (ignoring `generated_at`) and
skips the write if nothing meaningful changed.
"""

import json
from pathlib import Path
from typing import Dict


def _spec_content_equal(old_spec: Dict, new_spec: Dict) -> bool:
    """Compare two spec dicts, ignoring the `generated_at` field."""
    def _without_generated_at(d: Dict) -> Dict:
        return {k: v for k, v in d.items() if k != 'generated_at'}

    return _without_generated_at(old_spec) == _without_generated_at(new_spec)


def write_spec_if_changed(
    spec_path: Path,
    spec: Dict,
    *,
    sort_keys: bool = True,
) -> bool:
    """Write spec.json only if the content has meaningfully changed.

    Compares `spec` against the existing file at `spec_path`, ignoring the
    `generated_at` timestamp.  If nothing else changed the file is left
    untouched so that `git diff` stays clean.

    Args:
        spec_path: Path to spec.json.
        spec: The new spec dict (must already include `generated_at`).
        sort_keys: Whether to sort JSON keys (default True, matching most
            generators).

    Returns:
        True if the file was written (new or changed), False if skipped.
    """
    spec_path = Path(spec_path)

    if spec_path.exists():
        try:
            with open(spec_path, 'r', encoding='utf-8') as f:
                old_spec = json.load(f)
            if _spec_content_equal(old_spec, spec):
                return False  # unchanged — skip write
        except (json.JSONDecodeError, OSError):
            pass  # corrupted or unreadable — overwrite

    spec_path.parent.mkdir(parents=True, exist_ok=True)
    with open(spec_path, 'w', encoding='utf-8') as f:
        json.dump(spec, f, indent=2, ensure_ascii=False, sort_keys=sort_keys)
    return True
