# Word Revision comparison experiment

Word Revision compares a prediction with a reference using Microsoft Word's
CompareDocuments API. It is an optional comparison baseline. The official
WordEval score uses Word Object.

Use Windows, Microsoft Word, Python, pywin32 and psutil. Run Word COM evaluations
sequentially. From the repository or an export containing experiments:

```powershell
python -m experiments.word_revision --wordeval PATH_TO_WORDEVAL --predictions PATH_TO_PREDICTIONS --output word_revision_results.json
```

The input directory uses the public track/case layout accepted by
`scripts/evaluate_wordeval.py`. The entry point selects only Word Revision and
uses the same package checks, isolated workers, normal-open protocol and file
ownership policy as the core evaluator. Repair is diagnostic; repaired content
is never used as a successful prediction. Unknown or environmental failures
remain errors. Zero revisions is this baseline's criterion and does not imply
that Word Object's task-specific requirements were checked.

Reports identify `normal-open-shared-runtime-v1` and hash the relevant sources.
This version includes opening and revision-enumeration fixes made after the
paper's comparison was run, so it is a baseline implementation rather than an
exact reproduction of the appendix numbers. The comparison samples, human
labels and results from the paper are not part of this repository.

The same method can also be selected from the main evaluator with
`python scripts/evaluate_wordeval.py --methods word_object word_revision`.

Core imports are lazy: installing or running Word Object does not import Word
Revision. The shared document modules never import either scoring method.
