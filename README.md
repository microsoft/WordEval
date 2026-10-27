# WordEval

**WordEval: Evaluating Word-Native Operation Fidelity in Document Editing**

WordEval evaluates document editing through Microsoft Word's native object
state. The benchmark contains **642 tasks: 546 Atomic and 96 Compositional**.
This repository includes the benchmark, its evaluator, and generation source.

## Install

Use Python 3.10 or later on Windows with desktop Microsoft Word installed and
activated. Word-backed evaluation requires a working interactive Office session.
Office version, fonts, language, and local configuration can affect results;
the evaluator records the environment in its report.

```powershell
python -m pip install -r requirements.txt
```

Run Word evaluation sequentially on each machine. The default evaluator uses
isolated workers with a 180-second timeout per case. Do not disable isolation
for a formal benchmark run.

## Run a document editing agent

Create the solver-facing input folders from the included benchmark:

```powershell
python scripts/export_wordeval_inputs.py --source wordeval --output wordeval_inputs
```

Each `wordeval_inputs/{track}/{case_id}/` contains `input.docx` and
`instruction.txt`; the exporter also writes `wordeval_inputs/TASK_DESCRIPTION.md`
with the task rules. Give the agent only those files, and save its answer as
`pred.docx` in the corresponding case folder. Do not expose oracle metadata or
reference documents to the agent.

## Evaluate predictions

```powershell
python scripts/evaluate_wordeval.py --wordeval wordeval --predictions wordeval_inputs --release-manifest version.json --output results.json
```

The release manifest verifies the bundled evaluator and benchmark before any
Word work. Reports retain pass, fail, missing, and error counts. The primary
score is passes divided by all expected cases, including missing predictions
in the denominator (546 Atomic, 96 Compositional, 642 overall).

The Word Revision comparison from the paper appendix is available as an
optional method: add `--methods word_object word_revision`. It is not part of
the primary score. See `experiments/word_revision/README.md`.

## Repository contents

| Path | Contents |
|---|---|
| `wordeval/` | Frozen benchmark: instructions, inputs, references, and oracle metadata |
| `evaluation/` | Word Object evaluator and isolated document runtime |
| `scripts/evaluate_wordeval.py` | Evaluation entry point |
| `experiments/word_revision/` | Optional Word Revision comparison experiment |
| `generator/`, `instructions/` | Benchmark construction source |
| `scripts/export_wordeval*.py` | Canonical benchmark and solver-input exporters |
| `version.json` | Frozen evaluator and benchmark version manifest |
| `DATA_README.md`, `DATA_CARD.md` | Dataset layout, intended use, and limitations |

The bundled benchmark and evaluator are the version used for the reported
NeurIPS 2026 results. `version.json` records their fingerprints, and the
evaluator verifies them before any Word work when `--release-manifest` is
given. Text files use LF line endings. Archive checksums are generated when
packaging a release and distributed separately from the source tree.

Regeneration is a development operation and can depend on the installed Office
environment; use the checked-in benchmark for comparison with the reported
results. For generation, use Python 3.11+ and install
`requirements-generation.txt`, then run `python generator/test_sets/run.py --help`.
The generator writes a
`benchmark/` tree, which `python scripts/export_wordeval.py` converts into the
`wordeval/` layout. Export to a separate directory so the frozen benchmark
remains intact.

The project page and blog are maintained separately. This repository is the
code and benchmark distribution.

## External software and command metadata

Microsoft Word and Windows are external dependencies and are not bundled.
The generator uses Microsoft's [Office Fluent UI command identifier catalog](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)
as construction metadata for Word commands.

## License

Code is licensed under [MIT](LICENSE). Benchmark data under `wordeval/` and
the derived `wordeval_inputs/` layout use
[CDLA-Permissive-2.0](DATA_LICENSE).

## Synthetic data disclosure

WordEval is synthetic benchmark data. The documents, task instructions,
reference documents, and evaluator metadata are generated programmatically from
Microsoft Word editing operations and project-owned generation scripts. They
are not generated from customer data, personal data, or third-party datasets,
and third-party data is not used to seed synthetic data generation.

## Privacy

See Microsoft's [Privacy & Cookies](https://go.microsoft.com/fwlink/?LinkId=521839)
statement.

## Contributing and security

See [CONTRIBUTING.md](CONTRIBUTING.md), the
[Code of Conduct](CODE_OF_CONDUCT.md), and [SECURITY.md](SECURITY.md).

## Citation

WordEval has been accepted at **NeurIPS 2026**. Please cite:

```bibtex
@inproceedings{park2026wordeval,
  title     = {WordEval: Evaluating Word-Native Operation Fidelity in Document Editing},
  author    = {Jinyoung Park and Yupan Huang and Shimin Zhang and Jiayu Ding and Yilin Jia and Yuzhong Zhao and Tengchao Lv and Wenshan Wu and Shaohan Huang and Nan Yang and Xiangyang Zhou and Dongdong Zhang and Li Dong and Lei Cui and Sun Mao and Xun Wang and Tao Ge and Huitian Jiao and Si-Qing Chen and Changick Kim and Furu Wei},
  booktitle = {Advances in Neural Information Processing Systems},
  year      = {2026}
}
```

Machine-readable paper metadata is provided in [CITATION.cff](CITATION.cff).
