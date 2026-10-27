"""Match filled values without collapsing controls that share a title or tag."""
from collections import defaultdict


def filled_value_mismatches(gold, pred):
    remaining = defaultdict(list)
    for control in pred:
        key = control.get("title") or control.get("tag") or ""
        remaining[key].append((control.get("text") or "").strip())
    mismatches = []
    for control in gold:
        key = control.get("title") or control.get("tag") or ""
        value = (control.get("text") or "").strip()
        if not key or not value or (value.startswith("[") and value.endswith("]")):
            continue
        if value in remaining[key]:
            remaining[key].remove(value)
        else:
            mismatches.append({"field": key, "gold": value, "pred_values": list(remaining[key])})
    return mismatches
