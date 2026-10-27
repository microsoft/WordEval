# WordEval Data Card

## Dataset Summary

WordEval is a benchmark for evaluating Word-native operation fidelity in agents
that edit Microsoft Word documents. Each case provides a natural-language
instruction and an input `.docx` document. The agent must produce an edited
`.docx` document.

The release contains two tracks:

- **Atomic:** 546 single-operation editing cases.
- **Compositional:** 96 multi-operation cases that require composing several
  Word editing operations.

The canonical dataset is stored under `wordeval/`. The `wordeval_inputs/`
layout is a task-input convenience view generated from `wordeval/` and is not
the canonical source of record.

## Dataset Structure

```text
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
```

`cases.jsonl` contains task-input records with the case id, instruction, and
input document path. `oracle.jsonl` contains evaluator-only records with
reference documents, evaluation specifications, provenance keys, and case
metadata. The oracle files should not be shown during task execution.

## Generation and Collection

WordEval is synthetically generated from Microsoft Word editing operations.
Input and reference documents are created programmatically using generation
scripts and Microsoft Word automation. Reference documents are not manually
authored user documents.

The benchmark uses synthetic names, dates, addresses, email addresses, company
names, document titles, and document content. These strings are fixed test
fixtures for document editing cases and do not identify real data subjects.

## Intended Use

WordEval is intended for:

- Evaluating Word document editing agents.
- Comparing editing accuracy across operation types.
- Studying atomic versus compositional document editing behavior.
- Testing evaluation methods for Office-style document artifacts.

It is not intended for training models on private documents, evaluating general
writing quality, or measuring user productivity.

## Evaluation

The primary evaluator is `word_object`, which compares Word document state via
Microsoft Word COM APIs. It requires Windows and Microsoft Word. COM-based
evaluation should be run sequentially because Word automation is stateful and
not reliably parallel-safe on a single desktop session.

The evaluator reports case-level pass/fail/missing/error counts. Missing
predictions are counted separately from evaluated failures.

## Known Limitations

- Evaluation requires Windows and Microsoft Word.
- Some Word COM behavior can depend on the installed Word version, language
  pack, fonts, and local Office configuration.
- The benchmark focuses on `.docx` editing accuracy, not end-user interaction
  latency or UI usability.
- Some cases exercise highly specific Word features such as galleries,
  content controls, equations, charts, SmartArt, and object effects.

## Synthetic Data Disclosure

WordEval is synthetic benchmark data. The documents, task instructions,
reference documents, and evaluator metadata are generated programmatically from
Microsoft Word editing operations and project-owned generation scripts. They
are not generated from customer data, personal data, or third-party datasets,
and third-party data is not used to seed synthetic data generation.

## Privacy and Anonymity

The dataset is synthetic. The release documents and DOCX metadata have been
checked for obvious author names, local filesystem paths, private URLs, and
server identifiers. Synthetic fixture values such as `Megan Bowen`,
`Contoso Ltd.`, and `jane.doe@example.com` are part of the benchmark content
and are not author identifiers.

## Licensing

Benchmark data under `wordeval/` and the derived `wordeval_inputs/` layout is
released under the Community Data License Agreement - Permissive - Version 2.0
(`CDLA-Permissive-2.0`).
