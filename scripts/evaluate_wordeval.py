"""Evaluate predictions against a WordEval release.

This script reads the public WordEval schema:

    wordeval/{track}/oracle.jsonl

and expects predictions in the agent-facing layout:

    wordeval_inputs/{track}/{case_id}/pred.docx

Available methods:
- word_object: default structured Word object model checks.
- word_revision: Word CompareDocuments check; select it with
  --methods word_revision, or run both with --methods word_object word_revision.

Both methods use Word COM and require Windows with Microsoft Word installed.

Scoring is case-level by default. Each public {track}/{case_id} is one scored
case per method. pass_rate is passed / expected * 100, where expected includes
passed, failed, error, and missing cases. Missing predictions are counted as
zero-score cases for pass_rate and also reported separately. Track and overall
summaries micro-average by summing case counts. The report also includes macro summaries
over source_command groups: atomic groups correspond to commands, and
compositional groups correspond to named compositional units.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable

EVALUATION_ROOT = os.environ.get("WCB_EVALUATION_ROOT")
if EVALUATION_ROOT:
    # An environment override silently swapping the evaluator tree defeats
    # snapshot pinning; make the override loud and verifiable.
    print(
        f"[evaluate_wordeval] WCB_EVALUATION_ROOT override in effect: "
        f"{EVALUATION_ROOT}",
        file=sys.stderr,
    )
SCRIPT_PATH = Path(__file__).resolve()
SCRIPT_DIR = SCRIPT_PATH.parent
if not EVALUATION_ROOT and (SCRIPT_DIR / "evaluation").exists():
    EVALUATION_ROOT = str(SCRIPT_DIR / "evaluation")
if EVALUATION_ROOT:
    evaluation_parent = str(Path(EVALUATION_ROOT).resolve().parent)
    if evaluation_parent not in sys.path:
        sys.path.insert(0, evaluation_parent)

PROJECT_ROOT = SCRIPT_DIR.parent
for parent in (SCRIPT_DIR, *SCRIPT_DIR.parents):
    if (parent / "web" / "data" / "leaderboard_versions.json").exists():
        PROJECT_ROOT = parent
        break
if str(PROJECT_ROOT) not in sys.path:
    sys.path.append(str(PROJECT_ROOT))

from evaluation.evaluator import Evaluator

# Reporting belongs to this runner, including when an explicitly selected older
# frozen evaluator has no provenance module. Always fingerprint the imported
# evaluator below, rather than accidentally identifying the helper's own tree.
_provenance_path = next((path for path in (
    SCRIPT_DIR / "evaluation/provenance.py", SCRIPT_DIR.parent / "evaluation/provenance.py"
) if path.is_file()), None)
if _provenance_path is None:
    raise RuntimeError("Public runner is missing its bundled provenance helper")
_provenance_spec = importlib.util.spec_from_file_location("_wordeval_report_provenance", _provenance_path)
_provenance = importlib.util.module_from_spec(_provenance_spec)
_provenance_spec.loader.exec_module(_provenance)
evaluator_identity = _provenance.evaluator_identity
file_identity = _provenance.file_identity
runtime_environment = _provenance.runtime_environment
validate_release_manifest = _provenance.validate_release_manifest

# Fail closed if the imported evaluation package is not the tree this script
# resolved (e.g. the snapshot's evaluation/ is missing and the import silently
# fell through to a live repo on sys.path).
if EVALUATION_ROOT:
    _expected_root = Path(EVALUATION_ROOT).resolve()
    _actual_root = Path(sys.modules["evaluation"].__file__).resolve().parent
    if _actual_root != _expected_root:
        sys.exit(
            f"[evaluate_wordeval] imported 'evaluation' from {_actual_root}, "
            f"expected {_expected_root}; refusing to evaluate with a "
            "mismatched evaluator tree"
        )


DEFAULT_TRACKS = ("atomic", "compositional")
DEFAULT_CASE_TIMEOUT_SECONDS = 180.0
INFRA_ERROR_RATE_THRESHOLD = 0.01
INFRA_FAILURE_EXIT_CODE = 2
INFRA_GATING_METHODS = frozenset({"word_object"})
SCORE_DEFINITION = {
    "unit": "case",
    "pass_rate": "passed / expected * 100",
    "expected": "passed + failed + error + missing",
    "evaluated": "passed + failed + error",
    "missing": "Missing pred.docx files count as zero-score cases in pass_rate and are reported separately.",
    "case_micro": "Track and overall methods summaries sum case counts per method.",
    "group_macro": (
        "Group macro summaries first compute pass_rate within each source_command group, "
        "then average group pass_rate values equally. Groups with only missing predictions "
        "are included with 0 pass_rate."
    ),
    "group_key": (
        "source_command. In the atomic track this is a Word command; in the compositional "
        "track this is a named compositional unit."
    ),
}


def discover_tracks(wordeval_dir: Path) -> list[str]:
    """Discover public tracks from oracle files, with a v1-compatible fallback."""
    if not wordeval_dir.exists():
        return list(DEFAULT_TRACKS)
    tracks = [
        path.name
        for path in sorted(wordeval_dir.iterdir())
        if path.is_dir() and (path / "oracle.jsonl").exists()
    ]
    return tracks or list(DEFAULT_TRACKS)


def require_unique(values: Iterable[str], label: str) -> list[str]:
    """Return selectors unchanged, rejecting duplicates that skew summaries."""
    result = list(values)
    seen = set()
    duplicates = []
    for value in result:
        if value in seen and value not in duplicates:
            duplicates.append(value)
        seen.add(value)
    if duplicates:
        raise ValueError(f"Duplicate {label}: {', '.join(duplicates)}")
    return result


def read_jsonl(path: Path) -> Iterable[Dict[str, Any]]:
    with path.open("r", encoding="utf-8") as f:
        for line_number, line in enumerate(f, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                yield json.loads(line)
            except json.JSONDecodeError as exc:
                raise ValueError(f"Invalid JSON on {path}:{line_number}: {exc}") from exc


def read_oracle(path: Path) -> list[Dict[str, Any]]:
    """Validate identities before scoring so rows cannot overwrite one another."""
    rows = list(read_jsonl(path))
    if not rows:
        raise ValueError(f"Oracle contains no cases: {path}")
    ids = []
    for index, row in enumerate(rows, 1):
        if not isinstance(row, dict):
            raise ValueError(f"Oracle row {index} in {path} must be an object")
        case_id = row.get("case_id")
        if (not isinstance(case_id, str)
                or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*", case_id)
                or case_id.endswith(".")
                or case_id.split(".")[0].upper() in {
                    "CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                    *(f"LPT{i}" for i in range(1, 10))}):
            raise ValueError(f"Invalid case_id in oracle row {index} of {path}: {case_id!r}")
        ids.append(case_id)
        for field in ("input_document", "reference_document"):
            if field not in row:
                continue
            value = row[field]
            if not isinstance(value, str) or not value or Path(value).is_absolute():
                raise ValueError(f"Invalid {field} in oracle row {index} of {path}")
            resolved = (path.parent / value).resolve()
            if not resolved.is_relative_to(path.parent.resolve()):
                raise ValueError(f"{field} escapes track directory in oracle row {index} of {path}")
    require_unique(ids, f"oracle case_ids in {path}")
    require_unique([case_id.casefold() for case_id in ids], f"case-insensitive oracle case_ids in {path}")
    return rows


def empty_summary() -> Dict[str, int]:
    return {"evaluated": 0, "passed": 0, "failed": 0, "missing": 0, "error": 0}


def update_summary(summary: Dict[str, int], status: str) -> None:
    if status == "passed":
        summary["evaluated"] += 1
        summary["passed"] += 1
    elif status == "failed":
        summary["evaluated"] += 1
        summary["failed"] += 1
    elif status == "missing":
        summary["missing"] += 1
    else:
        summary["evaluated"] += 1
        summary["error"] += 1


def finalize_summary(summary: Dict[str, int]) -> Dict[str, Any]:
    evaluated = summary["evaluated"]
    expected = evaluated + summary["missing"]
    result = dict(summary)
    result["expected"] = expected
    result["pass_rate"] = round((summary["passed"] / expected) * 100, 1) if expected else 0.0
    result["evaluated_pass_rate"] = round((summary["passed"] / evaluated) * 100, 1) if evaluated else 0.0
    return result


def build_infrastructure_summary(
    method_summaries: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    """Build a non-scoring evaluator-infrastructure health summary.

    Infrastructure error rate is measured over cases that reached an
    evaluator method. Missing predictions remain in the score denominator,
    but cannot themselves produce evaluator errors and therefore do not
    dilute this operational health signal.
    """
    methods = {}
    for method, summary in method_summaries.items():
        error_count = int(summary.get("error", 0))
        evaluated_count = int(summary.get("evaluated", 0))
        missing_count = int(summary.get("missing", 0))
        rate = error_count / evaluated_count if evaluated_count else 0.0
        methods[method] = {
            "error_count": error_count,
            "evaluated_count": evaluated_count,
            "missing_count": missing_count,
            "infra_error_rate": rate,
            "threshold": INFRA_ERROR_RATE_THRESHOLD,
            "exceeded": rate > INFRA_ERROR_RATE_THRESHOLD,
            "gating": method in INFRA_GATING_METHODS,
        }

    return {
        "rate_definition": "error_count / evaluated_count",
        "rate_unit": "ratio",
        "gate_exceeded": any(
            method_summary["gating"] and method_summary["exceeded"]
            for method_summary in methods.values()
        ),
        "methods": methods,
    }


def report_exit_code(report: Dict[str, Any], fail_on_failure: bool) -> int:
    """Return the CLI exit code, with the infrastructure gate taking priority."""
    if report.get("infrastructure", {}).get("gate_exceeded", False):
        return INFRA_FAILURE_EXIT_CODE
    if fail_on_failure:
        failed = any(
            summary["failed"] or summary["error"]
            for summary in report["methods"].values()
        )
        return 1 if failed else 0
    return 0


def finalize_group_summary(group_summaries: Dict[str, Dict[str, int]]) -> Dict[str, Any]:
    groups = {
        group_name: finalize_summary(summary)
        for group_name, summary in sorted(group_summaries.items())
    }
    evaluated_groups = [
        summary for summary in groups.values()
        if summary["evaluated"] > 0
    ]
    scoring_groups = list(groups.values())
    macro_pass_rate = (
        round(sum(summary["pass_rate"] for summary in scoring_groups) / len(scoring_groups), 1)
        if scoring_groups
        else 0.0
    )
    return {
        "macro_pass_rate": macro_pass_rate,
        "group_count": len(groups),
        "evaluated_group_count": len(evaluated_groups),
        "missing_group_count": len(groups) - len(evaluated_groups),
        "groups": groups,
    }


def update_method_status(
    summaries: Dict[str, Dict[str, int]],
    group_summaries: Dict[str, Dict[str, Dict[str, int]]],
    methods: list[str],
    group_name: str,
    statuses: Dict[str, str],
) -> None:
    for method in methods:
        status = statuses.get(method, "error")
        update_summary(summaries[method], status)
        update_summary(group_summaries[method][group_name], status)


def evaluate_track(
    evaluator: Evaluator,
    methods: list[str],
    wordeval_dir: Path,
    predictions_dir: Path,
    track: str,
    *,
    oracle_rows: list[Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    track_dir = wordeval_dir / track
    oracle_path = track_dir / "oracle.jsonl"
    if not oracle_path.exists():
        raise FileNotFoundError(f"Missing oracle: {oracle_path}")

    summaries = {method: empty_summary() for method in methods}
    group_summaries = {method: defaultdict(empty_summary) for method in methods}
    results = {}

    for row in read_oracle(oracle_path) if oracle_rows is None else oracle_rows:
        case_id = row["case_id"]
        group_name = row.get("source_command") or "ungrouped"
        pred_path = predictions_dir / track / case_id / "pred.docx"

        case_result = {
            "instruction": row.get("instruction", ""),
            "source_command": row.get("source_command"),
            "source_variant": row.get("source_variant"),
            "source_test_set": row.get("source_test_set"),
            "source_key": row.get("source_key"),
            "methods": {},
            "artifacts": {"prediction": file_identity(pred_path)},
        }
        for role, field in (("input", "input_document"), ("reference", "reference_document")):
            if field in row:
                case_result["artifacts"][role] = file_identity(track_dir / row[field])

        if not pred_path.exists():
            statuses = {}
            for method in methods:
                case_result["methods"][method] = {"status": "missing"}
                statuses[method] = "missing"
            update_method_status(summaries, group_summaries, methods, group_name, statuses)
            results[case_id] = case_result
            continue

        spec = dict(row.get("evaluation_spec") or {})
        spec["idMso"] = row.get("source_command")
        spec["variation"] = row.get("source_variant")

        method_results = evaluator.evaluate_case(
            pred_path=pred_path,
            gold_path=track_dir / row["reference_document"],
            doc_in_path=track_dir / row["input_document"],
            spec=spec,
        )
        case_result["methods"].update(method_results)

        statuses = {
            method: method_results.get(method, {}).get("status", "error")
            for method in methods
        }
        update_method_status(summaries, group_summaries, methods, group_name, statuses)

        results[case_id] = case_result

    return {
        "track": track,
        "group_key": "source_command",
        "methods": {method: finalize_summary(summary) for method, summary in summaries.items()},
        "groups": {
            "group_key": "source_command",
            "methods": {
                method: finalize_group_summary(group_summaries[method])
                for method in methods
            },
        },
        "results": results,
    }


def evaluate_wordeval(
    wordeval_dir: Path,
    predictions_dir: Path,
    tracks: Iterable[str],
    methods: list[str],
    case_timeout_seconds: float | None = DEFAULT_CASE_TIMEOUT_SECONDS,
    release_manifest: Path | None = None,
) -> Dict[str, Any]:
    tracks = require_unique(tracks, "tracks")
    methods = require_unique(methods, "methods")
    require_unique([track.casefold() for track in tracks], "case-insensitive tracks")
    if not tracks or not methods:
        raise ValueError("At least one track and one method must be selected")
    if any(not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]*", track) for track in tracks):
        raise ValueError("Track selectors must name immediate track directories")
    # Validate every selected track before starting Word or evaluating any case.
    oracles = {track: read_oracle(wordeval_dir / track / "oracle.jsonl") for track in tracks}
    runtime_root = Path(sys.modules["evaluation"].__file__).resolve().parent
    release_integrity = {"verified": False, "reason": "no_release_manifest_selected"}
    if release_manifest is not None:
        release_integrity = validate_release_manifest(
            release_manifest, runtime_root=runtime_root, runner_path=SCRIPT_PATH,
            wordeval_dir=wordeval_dir, selected_tracks=tracks, read_oracle=read_oracle,
            predictions_dir=predictions_dir)
    provenance = {
        "evaluator": evaluator_identity(runtime_root),
        "runner": file_identity(SCRIPT_PATH),
        "environment": runtime_environment(),
        "asset_version_manifest": file_identity(wordeval_dir.parent / "version.json"),
        "oracles": {track: file_identity(wordeval_dir / track / "oracle.jsonl") for track in tracks},
        "case_timeout_seconds": case_timeout_seconds,
        "release_integrity": release_integrity,
    }
    evaluator = Evaluator(
        methods=methods,
        case_timeout_seconds=case_timeout_seconds,
    )
    track_reports = {}
    overall = {method: empty_summary() for method in methods}
    overall_groups = {method: defaultdict(empty_summary) for method in methods}

    for track in tracks:
        report = evaluate_track(evaluator, methods, wordeval_dir, predictions_dir, track,
                                oracle_rows=oracles[track])
        track_reports[track] = report
        for method, summary in report["methods"].items():
            for key in ("evaluated", "passed", "failed", "missing", "error"):
                overall[method][key] += summary[key]
        for method in methods:
            track_groups = report["groups"]["methods"][method]["groups"]
            for group_name, group_summary in track_groups.items():
                overall_group_name = f"{track}/{group_name}"
                for key in ("evaluated", "passed", "failed", "missing", "error"):
                    overall_groups[method][overall_group_name][key] += group_summary[key]

    method_summaries = {
        method: finalize_summary(summary)
        for method, summary in overall.items()
    }
    return {
        "wordeval": str(wordeval_dir),
        "predictions": str(predictions_dir),
        "score_definition": SCORE_DEFINITION,
        "provenance": provenance,
        "methods": method_summaries,
        "infrastructure": build_infrastructure_summary(method_summaries),
        "groups": {
            "group_key": "track/source_command",
            "methods": {
                method: finalize_group_summary(overall_groups[method])
                for method in methods
            },
        },
        "tracks": track_reports,
    }


def main(argv=None, *, default_methods=None, allowed_methods=None, experiment_metadata=None) -> int:
    parser = argparse.ArgumentParser(
        description="Evaluate WordEval predictions.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Methods:
  word_object    Default structured Word object model checks.
  word_revision  Word CompareDocuments checks. Use --methods word_revision.

Scoring:
  Each case_id counts once per method. pass_rate = passed / expected * 100.
  expected includes passed, failed, error, and missing. Missing pred.docx files are
  counted as zero-score cases and also reported separately. Track and overall methods
  summaries are case-level micro averages. The report also includes group macro
  summaries over source_command: atomic commands and named compositional units.

Infrastructure gate:
  Reports error_count / evaluated_count for every method. Only word_object is
  gating by default; a rate strictly above 1% exits with status 2.
""",
    )
    parser.add_argument("--wordeval", default="wordeval", help="WordEval release directory.")
    parser.add_argument("--release-manifest", type=Path,
                        help="Verify immutable release runtime, runner, assets and complete oracle against version.json before scoring.")
    parser.add_argument(
        "--predictions",
        default="wordeval_inputs",
        help="Prediction directory with {track}/{case_id}/pred.docx files.",
    )
    parser.add_argument(
        "--track",
        action="append",
        help="Track to evaluate. Can be specified multiple times. Defaults to all tracks with oracle.jsonl.",
    )
    parser.add_argument(
        "--methods",
        nargs="+",
        default=default_methods or ["word_object"],
        choices=allowed_methods,
        help=f"Evaluation methods to run (default: {', '.join(default_methods or ['word_object'])}).",
    )
    parser.add_argument(
        "--case-timeout-seconds",
        type=float,
        default=DEFAULT_CASE_TIMEOUT_SECONDS,
        help=(
            "Run Word-backed methods in an isolated worker with this per-case "
            "timeout (default: 180; set to 0 only for non-COM debugging)."
        ),
    )
    parser.add_argument(
        "--output",
        default="wordeval_results.json",
        help="Output report path (default: wordeval_results.json).",
    )
    parser.add_argument(
        "--fail-on-failure",
        action="store_true",
        help=(
            "Exit 1 when evaluated cases fail or error. The default word_object "
            "infrastructure gate exits 2 when its error rate exceeds 1%%, and "
            "takes precedence."
        ),
    )
    args = parser.parse_args(argv)

    wordeval_dir = Path(args.wordeval)
    try:
        tracks = require_unique(
            args.track or discover_tracks(wordeval_dir),
            "tracks",
        )
        args.methods = require_unique(args.methods, "methods")
    except ValueError as exc:
        parser.error(str(exc))
    missing_tracks = [track for track in tracks if not (wordeval_dir / track / "oracle.jsonl").exists()]
    if missing_tracks:
        parser.error(f"Track oracle not found: {', '.join(missing_tracks)}")

    report = evaluate_wordeval(
        wordeval_dir=wordeval_dir,
        predictions_dir=Path(args.predictions),
        tracks=tracks,
        methods=args.methods,
        case_timeout_seconds=args.case_timeout_seconds,
        release_manifest=args.release_manifest,
    )
    if experiment_metadata is not None:
        report['experiment'] = experiment_metadata

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False, sort_keys=True)

    print(json.dumps(report["methods"], indent=2, sort_keys=True))
    print(f"Report saved to: {output_path}")

    return report_exit_code(report, args.fail_on_failure)


if __name__ == "__main__":
    raise SystemExit(main())
