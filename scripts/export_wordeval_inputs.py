"""Export solver-facing case folders from the WordEval public dataset.

Input:
    wordeval/
      atomic/cases.jsonl
      compositional/cases.jsonl

Output:
    wordeval_inputs/
      atomic/{case_id}/
        input.docx
        instruction.txt
      compositional/{case_id}/
        input.docx
        instruction.txt

The exporter intentionally reads only cases.jsonl, never oracle.jsonl,
references, or metadata. This keeps solver-facing packages free of answers and
evaluator metadata.
"""

from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable


DEFAULT_TRACKS = ("atomic", "compositional")

TASK_DESCRIPTION = """# WordEval Task Description

## Overview

Each task asks you to complete one WordEval case by editing a Microsoft Word
document according to a natural language instruction.

## Files

Each case folder contains:

```text
input.docx        # Input Word document for the case
instruction.txt  # Task instruction
pred.docx         # Your output, to be created
method.md         # Optional description of your approach
```

Some benchmark runners may request `method.md` for reproducibility. It is not
used by the automatic scorer.

## Task

For each case folder:

1. Read `instruction.txt`.
2. Open `input.docx`.
3. Apply the requested edit.
4. Save the edited document as `pred.docx` in the same folder.

You may use any approach.
"""

README = """# WordEval Inputs

This package contains the public, agent-facing WordEval cases. It intentionally
contains only inputs and instructions; oracle metadata and reference documents
are not included.

## Case Layout

Each case is stored as:

```text
{track}/{case_id}/
  input.docx
  instruction.txt
```

The public tracks are:

- `atomic`
- `compositional`

## What To Produce

For every case you run, create:

```text
{track}/{case_id}/pred.docx
```

You may also include:

```text
{track}/{case_id}/method.md
```

`method.md` is optional and is not used by the automatic scorer.

## Submission Format

Submit a zip with one of these layouts:

```text
{model_name}.zip
  atomic/{case_id}/pred.docx
  compositional/{case_id}/pred.docx
```

or:

```text
{model_name}.zip
  {model_name}/atomic/{case_id}/pred.docx
  {model_name}/compositional/{case_id}/pred.docx
```

The evaluator accepts full or partial submissions. Missing cases are reported as
missing and are not counted as passed.

Model names may contain only letters, numbers, `.`, `-`, and `_`, must be at
most 64 characters, and cannot start or end with `.`.
"""


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


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def copy_required(src: Path, dest: Path) -> None:
    if not src.exists():
        raise FileNotFoundError(f"Missing source file: {src}")
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)


def export_track(source_dir: Path, output_dir: Path, track: str) -> Dict[str, Any]:
    track_dir = source_dir / track
    cases_path = track_dir / "cases.jsonl"
    if not cases_path.exists():
        raise FileNotFoundError(f"Missing cases file: {cases_path}")

    output_track_dir = output_dir / track
    output_track_dir.mkdir(parents=True, exist_ok=True)

    case_count = 0
    manifest_cases = []

    for case in read_jsonl(cases_path):
        case_id = case.get("case_id")
        if not case_id:
            raise ValueError(f"Case without case_id in {cases_path}")

        case_output_dir = output_track_dir / case_id
        case_output_dir.mkdir(parents=True, exist_ok=True)

        doc_rel = case.get("input_document")
        if not doc_rel:
            raise ValueError(f"Case {case_id} is missing input_document")
        copy_required(track_dir / doc_rel, case_output_dir / "input.docx")

        write_text(case_output_dir / "instruction.txt", case.get("instruction", ""))

        manifest_entry = {
            "case_id": case_id,
            "path": f"{track}/{case_id}",
        }

        manifest_cases.append(manifest_entry)
        case_count += 1

    return {
        "track": track,
        "case_count": case_count,
        "cases": manifest_cases,
    }


def export_wordeval_inputs(
    source_dir: Path,
    output_dir: Path,
    tracks: Iterable[str],
    clean: bool = False,
) -> Dict[str, Any]:
    source_dir = source_dir.resolve()
    output_dir = output_dir.resolve()

    if not source_dir.exists():
        raise FileNotFoundError(f"WordEval source directory does not exist: {source_dir}")

    available_tracks = {
        path.name
        for path in source_dir.iterdir()
        if path.is_dir() and (path / "cases.jsonl").exists()
    }
    selected_tracks = tuple(tracks)
    unknown = sorted(set(selected_tracks) - available_tracks)
    if unknown:
        raise ValueError(f"Unknown track(s): {', '.join(unknown)}")

    if clean and output_dir.exists():
        shutil.rmtree(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    track_manifests = {}
    for track in selected_tracks:
        track_manifests[track] = export_track(source_dir, output_dir, track)

    public_manifest = {
        "name": "WordEval agent-facing inputs",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "tracks": {
            track: {
                "case_count": data["case_count"],
            }
            for track, data in track_manifests.items()
        },
        "total_case_count": sum(data["case_count"] for data in track_manifests.values()),
    }
    with (output_dir / "manifest.json").open("w", encoding="utf-8") as f:
        json.dump(public_manifest, f, indent=2, ensure_ascii=False, sort_keys=True)
    write_text(output_dir / "README.md", README)
    write_text(output_dir / "TASK_DESCRIPTION.md", TASK_DESCRIPTION)

    for track, data in track_manifests.items():
        manifest_path = output_dir / track / "manifest.json"
        with manifest_path.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False, sort_keys=True)

    return public_manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="Export agent-facing WordEval case folders.")
    parser.add_argument(
        "--source",
        default="wordeval",
        help="Input WordEval directory (default: ./wordeval)",
    )
    parser.add_argument(
        "--output",
        default="wordeval_inputs",
        help="Output agent-facing directory (default: ./wordeval_inputs)",
    )
    parser.add_argument(
        "--track",
        action="append",
        help="Track to export. Can be specified multiple times. Defaults to all tracks with cases.jsonl.",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Remove the output directory before exporting.",
    )
    args = parser.parse_args()

    source_dir = Path(args.source)
    discovered_tracks = []
    if source_dir.exists():
        discovered_tracks = [
            path.name
            for path in sorted(source_dir.iterdir())
            if path.is_dir() and (path / "cases.jsonl").exists()
        ]
    tracks = args.track or discovered_tracks or list(DEFAULT_TRACKS)
    manifest = export_wordeval_inputs(
        source_dir=source_dir,
        output_dir=Path(args.output),
        tracks=tracks,
        clean=args.clean,
    )
    print(json.dumps(manifest, indent=2, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
