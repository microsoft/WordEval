"""Run one verification method in a disposable process."""

from __future__ import annotations

import json
import sys
from pathlib import Path

from .methods import get_method


def main() -> None:
    # stdin must decode the parent's UTF-8 payload; without this the Windows
    # locale codec (cp1252) mojibakes non-ASCII spec values and paths, which
    # mis-scores cases instead of erroring.
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    payload = json.load(sys.stdin)
    expected_root = payload.get("evaluation_root")
    if expected_root:
        import evaluation

        actual_root = Path(evaluation.__file__).resolve().parent
        if Path(expected_root).resolve() != actual_root:
            raise SystemExit(
                "isolated worker imported 'evaluation' from "
                f"{actual_root}, expected {expected_root}; refusing to "
                "verify with a mismatched evaluator tree"
            )
    method = None
    if payload.get("action") == "probe_word_openability":
        from .document_io.open_gate import normal_open_gate
        from .methods.base import VerificationResult
        result = VerificationResult.from_legacy(normal_open_gate(
            Path(payload["pred_path"]),
            Path(payload["gold_path"]),
        ))
    else:
        method = get_method(payload["method"])
        result = method.verify(
            Path(payload["pred_path"]),
            Path(payload["gold_path"]),
            Path(payload["doc_in_path"]) if payload.get("doc_in_path") else None,
            payload.get("spec"),
        )
    print(json.dumps(result.to_dict(), ensure_ascii=False), flush=True)
    cleanup = getattr(method, "cleanup", None)
    if cleanup:
        cleanup()


if __name__ == "__main__":
    main()
