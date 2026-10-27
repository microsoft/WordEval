"""Method-independent normal-open evidence and role attribution."""
from pathlib import Path
from .word import WordSession


def _result(passed, message, details=None, error=None):
    return {"passed": passed, "message": message, "details": details or {}, "error": error}


def normal_open_gate(pred_path, gold_path, *, probe=None):
    """Observe healthy reference and prediction independently; never score repairs."""
    probe = probe or WordSession.probe_clean_openability
    gold_evidence = probe(str(gold_path))
    if gold_evidence.get("state") != "normal_opened":
        error = "Gold failed the independent normal-open gate"
        return _result(
            passed=False,
            message=f"ERROR: {error}",
            details={
                "open_failure_owner": "gold",
                "normal_open_evidence": {"gold": gold_evidence},
            },
            error=error,
        )

    pred_evidence = probe(str(pred_path))
    evidence = {
        "gold": gold_evidence,
        "prediction": pred_evidence,
    }
    pred_state = pred_evidence.get("state")
    if pred_state == "normal_opened":
        return _result(
            passed=True,
            message="PASS: Prediction and gold open normally in clean Word",
            details={"normal_open_evidence": evidence},
        )
    if pred_state in {"repair_required", "unopenable"}:
        repair_required = pred_state == "repair_required"
        return _result(
            passed=False,
            message=(
                "INVALID OUTPUT: pred.docx requires Word repair"
                if repair_required
                else "INVALID OUTPUT: pred.docx is reproducibly unopenable"
            ),
            details={
                "reason": "invalid_output",
                "detected_format": (
                    "word_repair_required" if repair_required else "word_unopenable"
                ),
                "open_failure_owner": "prediction",
                "normal_open_evidence": evidence,
            },
            error=None,
        )

    error = "Prediction normal-open probe could not establish ownership"
    return _result(
        passed=False,
        message=f"ERROR: {error}",
        details={
            "open_failure_owner": "environment_or_unknown",
            "normal_open_evidence": evidence,
        },
        error=error,
    )

