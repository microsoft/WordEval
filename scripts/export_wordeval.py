"""Export the public-facing WordEval dataset layout.

This exporter converts the generated ``benchmark/`` tree plus ``instructions/``
into a clean release-oriented layout:

    wordeval/
      atomic/
        cases.jsonl
        oracle.jsonl
        inputs/
        references/
        metadata/
      compositional/
        cases.jsonl
        oracle.jsonl
        inputs/
        references/
        metadata/
      metadata/

The source ``wordeval/`` format remains the internal/legacy evaluator
format and is intentionally not used by this exporter.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from generator.export import (
    VARIANT_BASED_TEST_SETS,
    VARIANT_FALLBACK_TEST_SETS,
    _build_docx_identity_context,
    _merge_docx_identity_context,
    get_instruction,
    load_instructions,
    load_instructions_from_variants,
    load_instructions_with_variant_fallback,
    load_test_set_summary,
    sanitize_docx_identity,
)


TRACK_ATOMIC = "atomic"
TRACK_COMPOSITIONAL = "compositional"


def stable_hash(value: str, length: int = 12) -> str:
    """Return a short stable hash for IDs."""
    return hashlib.sha256(value.encode("utf-8")).hexdigest()[:length]


def file_hash(path: Path, length: int = 12) -> str:
    """Return a short hash of file bytes."""
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()[:length]


def json_dump_line(f, record: Dict[str, Any]) -> None:
    f.write(json.dumps(record, ensure_ascii=False, sort_keys=True))
    f.write("\n")


def load_public_instructions(instructions_dir: Path, test_set: str) -> Optional[Dict[str, Any]]:
    if test_set in VARIANT_BASED_TEST_SETS:
        return load_instructions_from_variants(instructions_dir, test_set)
    if test_set in VARIANT_FALLBACK_TEST_SETS:
        return load_instructions_with_variant_fallback(instructions_dir, test_set)
    return load_instructions(instructions_dir, test_set)


def iter_benchmark_cases(
    benchmark_dir: Path,
    instructions_dir: Path,
    include_no_effect: bool = False,
) -> Iterable[Dict[str, Any]]:
    valid_statuses = {"xml_effect_execution"}
    if include_no_effect:
        valid_statuses.add("no_xml_effect_execution")

    for test_set_dir in sorted(benchmark_dir.iterdir()):
        if not test_set_dir.is_dir():
            continue

        test_set = test_set_dir.name
        summary = load_test_set_summary(test_set_dir)
        if not summary:
            continue

        instructions = load_public_instructions(instructions_dir, test_set)
        if not instructions:
            raise ValueError(f"Could not load instructions for {test_set}")

        for command, variants in sorted((summary.get("test_cases") or {}).items()):
            for variant, status_entry in sorted((variants or {}).items()):
                if status_entry.get("status") not in valid_statuses:
                    continue

                variation_dir = test_set_dir / command / variant
                spec_path = variation_dir / "spec.json"
                if not spec_path.exists():
                    raise FileNotFoundError(f"Missing spec.json: {spec_path}")

                with spec_path.open("r", encoding="utf-8") as f:
                    full_spec = json.load(f)

                files = full_spec.get("files") or {}
                doc_in_path = (variation_dir / files.get("doc_in", "doc_in.docx")).resolve()
                gold_path = (variation_dir / files.get("gold", "gold.docx")).resolve()
                if not doc_in_path.exists():
                    raise FileNotFoundError(f"Missing doc_in: {doc_in_path}")
                if not gold_path.exists():
                    raise FileNotFoundError(f"Missing gold: {gold_path}")

                instruction = get_instruction(instructions, command, variant)
                entry = {
                    "instruction": instruction,
                    "spec": {
                        "verification": full_spec.get("verification"),
                    },
                }
                if "behavior" in full_spec:
                    entry["spec"]["behavior"] = full_spec["behavior"]
                if "precondition" in full_spec:
                    entry["spec"]["precondition"] = full_spec["precondition"]
                if "execution" in full_spec:
                    # The note verifiers read execution.insert_text / custom_text
                    # / custom_mark / insert_para_index / position to derive the
                    # instruction-owned note contract.  Without it the verifier
                    # falls back to an unrelated 'See reference.' default, which
                    # errors against gold's real note text (footnote/endnote
                    # standard_insert).  Carry it through so the exported spec
                    # matches what the verifier expects.
                    entry["spec"]["execution"] = full_spec["execution"]
                if "composition" in full_spec:
                    composition = full_spec.get("composition") or {}
                    if "commands" in composition:
                        raise ValueError(
                            f"{spec_path} contains legacy composition.commands; "
                            "regenerate this composition test set before WordEval export."
                        )
                    if not composition.get("atomic_intents"):
                        raise ValueError(
                            f"{spec_path} is a composition spec without "
                            "composition.atomic_intents; regenerate or fix the "
                            "composition generator before WordEval export."
                        )
                    entry["spec"]["composition"] = composition
                input_identity_context = _build_docx_identity_context(
                    full_spec, command, artifact_role="input")
                gold_identity_context = _build_docx_identity_context(
                    full_spec, command, artifact_role="gold")

                yield {
                    "test_set": test_set,
                    "test_set_dir": test_set_dir,
                    "command": command,
                    "variant": variant,
                    "entry": entry,
                    "doc_in_path": doc_in_path,
                    "gold_path": gold_path,
                    "input_identity_context": input_identity_context,
                    "gold_identity_context": gold_identity_context,
                }


def case_track(test_set: str, entry: Dict[str, Any]) -> str:
    spec = entry.get("spec") or {}
    if "composition" in spec or "composition" in test_set:
        return TRACK_COMPOSITIONAL
    return TRACK_ATOMIC


def behavior_from_entry(entry: Dict[str, Any]) -> Optional[str]:
    spec = entry.get("spec") or {}
    behavior = spec.get("behavior")
    return behavior if isinstance(behavior, str) else None


def source_key_for(track: str, command: str, variant: str) -> str:
    # Atomic commands use idMso/variant. Compositional commands use the composition
    # family key/variant. Both are intentionally independent of source test set.
    return f"{command}/{variant}"


def case_id_for(track: str, source_key: str) -> str:
    prefix = "case-atomic" if track == TRACK_ATOMIC else "case-compositional"
    return f"{prefix}-{stable_hash(f'{track}:{source_key}')}"


def document_id_for(path: Path) -> str:
    return f"input-{file_hash(path)}"


def resolve_source_path(test_set_dir: Path, rel_path: str) -> Path:
    path = (test_set_dir / rel_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"Missing source file: {path}")
    return path


def copy_once(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        return
    shutil.copy2(src, dest)


def write_track_metadata(
    track_dir: Path,
    track: str,
    records: list[Dict[str, Any]],
    source_set_counts: Counter,
    behavior_counts: Counter,
) -> None:
    metadata_dir = track_dir / "metadata"
    metadata_dir.mkdir(parents=True, exist_ok=True)

    manifest = {
        "track": track,
        "case_count": len(records),
        "input_document_count": len({record["input_id"] for record in records}),
        "source_set_counts": dict(sorted(source_set_counts.items())),
        "behavior_counts": dict(sorted(behavior_counts.items())),
        "generated_at": datetime.now(timezone.utc).isoformat(),
    }
    with (metadata_dir / "manifest.json").open("w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False, sort_keys=True)

    taxonomy_rows = []
    for record in records:
        row = {
            "case_id": record["case_id"],
            "source_key": record["source_key"],
            "source_test_set": record["source_test_set"],
            "source_command": record["source_command"],
            "source_variant": record["source_variant"],
        }
        if record.get("behavior"):
            row["behavior"] = record["behavior"]
        if record.get("composition"):
            composition = record["composition"]
            row["composition_type"] = composition.get("type")
            row["operation_count"] = composition.get("operation_count")
            row["atomic_commands"] = [
                intent.get("atomic_case")
                for intent in composition.get("atomic_intents", []) or []
                if intent.get("atomic_case")
            ]
        taxonomy_rows.append(row)

    with (metadata_dir / "taxonomy.json").open("w", encoding="utf-8") as f:
        json.dump(taxonomy_rows, f, indent=2, ensure_ascii=False, sort_keys=True)


def export_wordeval(
    benchmark_dir: Path,
    output_dir: Path,
    instructions_dir: Optional[Path] = None,
    include_no_effect: bool = False,
    clean: bool = False,
) -> Dict[str, Any]:
    benchmark_dir = benchmark_dir.resolve()
    output_dir = output_dir.resolve()

    if not benchmark_dir.exists():
        raise FileNotFoundError(f"Benchmark directory does not exist: {benchmark_dir}")
    if instructions_dir is None:
        instructions_dir = benchmark_dir.parent / "instructions"
    instructions_dir = instructions_dir.resolve()
    if not instructions_dir.exists():
        raise FileNotFoundError(f"Instructions directory does not exist: {instructions_dir}")

    if clean and output_dir.exists():
        shutil.rmtree(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    track_records: Dict[str, list[Dict[str, Any]]] = {
        TRACK_ATOMIC: [],
        TRACK_COMPOSITIONAL: [],
    }
    track_source_counts: Dict[str, Counter] = {
        TRACK_ATOMIC: Counter(),
        TRACK_COMPOSITIONAL: Counter(),
    }
    track_behavior_counts: Dict[str, Counter] = {
        TRACK_ATOMIC: Counter(),
        TRACK_COMPOSITIONAL: Counter(),
    }
    case_ids_seen: Dict[str, str] = {}
    input_ids_seen: Dict[tuple[str, str], Path] = {}

    cases_files = {}
    oracle_files = {}
    try:
        for track in (TRACK_ATOMIC, TRACK_COMPOSITIONAL):
            track_dir = output_dir / track
            track_dir.mkdir(parents=True, exist_ok=True)
            (track_dir / "inputs").mkdir(exist_ok=True)
            (track_dir / "references").mkdir(exist_ok=True)
            cases_files[track] = (track_dir / "cases.jsonl").open("w", encoding="utf-8")
            oracle_files[track] = (track_dir / "oracle.jsonl").open("w", encoding="utf-8")

        cases = list(iter_benchmark_cases(benchmark_dir, instructions_dir, include_no_effect=include_no_effect))
        input_contexts: Dict[Path, Dict[str, Any]] = {}
        for case in cases:
            doc_in_path = case["doc_in_path"]
            input_contexts[doc_in_path] = _merge_docx_identity_context(
                input_contexts.get(doc_in_path),
                case.get("input_identity_context") or {},
            )

        for case in cases:
            test_set = case["test_set"]
            command = case["command"]
            variant = case["variant"]
            entry = case["entry"]
            track = case_track(test_set, entry)
            source_key = source_key_for(track, command, variant)
            case_id = case_id_for(track, source_key)
            if case_id in case_ids_seen:
                previous = case_ids_seen[case_id]
                raise ValueError(
                    f"Case ID collision for {case_id}: {previous} and {test_set}/{source_key}"
                )
            case_ids_seen[case_id] = f"{test_set}/{source_key}"

            doc_in_path = case["doc_in_path"]
            reference_path = case["gold_path"]
            input_id = document_id_for(doc_in_path)
            input_dest = output_dir / track / "inputs" / f"{input_id}.docx"
            if not input_dest.exists():
                sanitize_docx_identity(doc_in_path, input_dest, input_contexts.get(doc_in_path))
            input_ids_seen[(track, input_id)] = input_dest

            reference_name = f"reference-{case_id}.docx"
            reference_dest = output_dir / track / "references" / reference_name
            sanitize_docx_identity(
                reference_path,
                reference_dest,
                case.get("gold_identity_context"),
            )

            case_record = {
                "case_id": case_id,
                "instruction": entry.get("instruction", ""),
                "input_document": f"inputs/{input_id}.docx",
            }
            oracle_record = {
                "case_id": case_id,
                "track": track,
                "source_key": source_key,
                "source_test_set": test_set,
                "source_command": command,
                "source_variant": variant,
                "input_id": input_id,
                "input_document": f"inputs/{input_id}.docx",
                "reference_document": f"references/{reference_name}",
                "instruction": entry.get("instruction", ""),
            }

            spec = entry.get("spec") or {}
            if spec:
                oracle_record["evaluation_spec"] = spec

            behavior = behavior_from_entry(entry)
            if behavior:
                oracle_record["behavior"] = behavior
                track_behavior_counts[track][behavior] += 1

            composition = spec.get("composition")
            if composition:
                oracle_record["composition"] = composition

            json_dump_line(cases_files[track], case_record)
            json_dump_line(oracle_files[track], oracle_record)

            track_records[track].append(oracle_record)
            track_source_counts[track][test_set] += 1
    finally:
        for f in cases_files.values():
            f.close()
        for f in oracle_files.values():
            f.close()

    for track in (TRACK_ATOMIC, TRACK_COMPOSITIONAL):
        write_track_metadata(
            output_dir / track,
            track,
            track_records[track],
            track_source_counts[track],
            track_behavior_counts[track],
        )

    top_metadata_dir = output_dir / "metadata"
    top_metadata_dir.mkdir(exist_ok=True)
    suite_manifest = {
        "name": "WordEval",
        "source_format": "benchmark",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "tracks": {
            track: {
                "case_count": len(track_records[track]),
                "input_document_count": len({record["input_id"] for record in track_records[track]}),
            }
            for track in (TRACK_ATOMIC, TRACK_COMPOSITIONAL)
        },
        "total_case_count": sum(len(records) for records in track_records.values()),
    }
    with (top_metadata_dir / "suite_manifest.json").open("w", encoding="utf-8") as f:
        json.dump(suite_manifest, f, indent=2, ensure_ascii=False, sort_keys=True)

    return suite_manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="Export the WordEval public-facing dataset layout.")
    parser.add_argument(
        "--benchmark",
        default="benchmark",
        help="Input benchmark directory (default: ./benchmark)",
    )
    parser.add_argument(
        "--instructions",
        default="instructions",
        help="Instructions directory (default: ./instructions)",
    )
    parser.add_argument(
        "--output",
        default="wordeval",
        help="Output WordEval directory (default: ./wordeval)",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Remove the output directory before exporting.",
    )
    parser.add_argument(
        "--include-no-effect",
        action="store_true",
        help="Also include no_xml_effect_execution cases from benchmark summaries.",
    )
    args = parser.parse_args()

    manifest = export_wordeval(
        benchmark_dir=Path(args.benchmark),
        output_dir=Path(args.output),
        instructions_dir=Path(args.instructions),
        include_no_effect=args.include_no_effect,
        clean=args.clean,
    )
    print(json.dumps(manifest, indent=2, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
