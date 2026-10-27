# WordEval: Evaluating Word-Native Operation Fidelity in Document Editing

WordEval is a benchmark for evaluating Word-native operation fidelity in
document editing agents. The dataset contains controlled Microsoft Word `.docx`
editing tasks with natural-language instructions, input documents, reference
documents, and evaluator metadata.

This README describes the benchmark data included in this repository. For intended use, limitations, synthetic data, privacy, and
Responsible AI notes, see `DATA_CARD.md`.

## Dataset Contents

The canonical benchmark is included under `wordeval/`:

```text
wordeval/
  atomic/
    cases.jsonl
    oracle.jsonl
    inputs/
    references/
    metadata/
      manifest.json
      taxonomy.json
  compositional/
    cases.jsonl
    oracle.jsonl
    inputs/
    references/
    metadata/
      manifest.json
      taxonomy.json
  metadata/
    suite_manifest.json
```

The release contains 642 cases:

- `atomic`: 546 single-operation editing cases with 112 input documents.
- `compositional`: 96 multi-operation cases with 38 input documents.

## Task Inputs

`cases.jsonl` contains the task inputs. Each row contains:

```json
{
  "case_id": "case-atomic-...",
  "instruction": "Apply the requested Word edit.",
  "input_document": "inputs/input-....docx"
}
```

Only `cases.jsonl` and the referenced input documents should be provided to the
entity completing the task, whether that is a human, an agent, or another
document-editing system.

## Evaluation Materials

`oracle.jsonl` contains evaluator-only metadata. Each row contains the hidden
reference document path, evaluation specification, source command metadata,
behavior category, and related provenance fields:

```json
{
  "behavior": "content_formatting",
  "case_id": "case-atomic-...",
  "evaluation_spec": {},
  "input_document": "inputs/input-....docx",
  "input_id": "input-...",
  "instruction": "Apply the requested Word edit.",
  "reference_document": "references/reference-case-atomic-....docx",
  "source_command": "Command",
  "source_key": "Command/variant",
  "source_test_set": "test_set_name",
  "source_variant": "variant",
  "track": "atomic"
}
```

Compositional cases may also include a `composition` object describing the
atomic editing intents that make up the case. Do not expose `oracle.jsonl`,
`references/`, or evaluator metadata during task execution.

`inputs/` contains the input `.docx` files. `references/` contains gold
reference `.docx` files used by the evaluator.

Each track's `metadata/manifest.json` records track-level counts, behavior
counts, and source test-set counts. Each track's `metadata/taxonomy.json`
records track-level taxonomy metadata. The top-level
`wordeval/metadata/suite_manifest.json` records suite-level counts across both
tracks.

Evaluation code is included in this repository alongside the task inputs,
reference documents, and evaluator-only metadata.

## Derived Task Package

The input exporter creates `wordeval_inputs/`, a derived task package
with one folder per case. This package is generated from `cases.jsonl` and
`inputs/` for benchmark runners that expect per-case folders. It is a
convenience layout, not the canonical source of record.

The derived package has this layout:

```text
wordeval_inputs/
  manifest.json
  TASK_DESCRIPTION.md
  atomic/
    manifest.json
    {case_id}/
      input.docx
      instruction.txt
  compositional/
    manifest.json
    {case_id}/
      input.docx
      instruction.txt
```

The derived `wordeval_inputs/` layout intentionally excludes `oracle.jsonl`,
reference documents, and evaluation metadata.

## Documentation and Metadata

- `DATA_CARD.md`: intended use, limitations, privacy, synthetic data, licensing,
  and Responsible AI notes.
- `DATA_LICENSE`: CDLA-Permissive-2.0 license for benchmark data under
  `wordeval/` and the derived `wordeval_inputs/` layout.

## Synthetic Data Disclosure

WordEval is synthetic benchmark data. The documents, task instructions,
reference documents, and evaluator metadata are generated programmatically from
Microsoft Word editing operations and project-owned generation scripts. They
are not generated from customer data, personal data, or third-party datasets,
and third-party data is not used to seed synthetic data generation.
