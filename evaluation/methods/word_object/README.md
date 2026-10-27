# Word object evaluator

Use `evaluation.Evaluator.evaluate_case()` for a fully checked result, including
package validation, error attribution and the configured process timeout.
`method.py` is the lower-level scoring adapter. The behavior facades
`object_effect.py`, `content_operations.py`, and `content_formatting.py` select
the command route and preserve the existing verifier API.

## Where to implement a change

- `families/<feature>/`: feature-specific comparisons, saved OOXML readers,
  COM readers, and verification rules. `verify.py` handles effects;
  `insert.py` handles insertion when the feature needs distinct contracts.
- `evaluation/document_io/word.py`: shared Word process ownership, document
  lifetime, repair probes, retries and cleanup. Local `session.py` and
  `errors.py` preserve previous import paths; `core.py` exposes readers through
  the shared `WordSession`.
- `ooxml/`: package, relationship, XML, and canonical text helpers.
- `preservation/`: shared checks for unrelated document content and objects.
- `policies/`: behavior-specific COM/saved-evidence combination policies.
- `composition*.py`: declared step ordering, step evidence, and shared
  composition checks. `effect_commands.py` owns effect command groups.

Families take explicit evidence callbacks. They do not import behavior facades
or take an entire verifier object. Deferred callbacks preserve document-open
order and declaration-before-observation errors. Pure comparisons accept data
and return the existing result objects without starting Word.

The behavior policies deliberately remain separate: their exception handling
and saved-evidence guardrails have different contracts. SmartArt composition
keeps its bounded cache ownership in the effect facade, restoring both caches
in `finally`. A saved target match alone does not establish preservation.

## Tests

Run `python -m pytest evaluation/tests -q -p no:cacheprovider`. Family tests
exercise deterministic evidence; routing tests and cross-family regressions
exercise compatibility entrypoints. Word COM checks run sequentially per guest.

For behavior-preserving changes, retain scenario inputs and expected results,
then compare the candidate with a frozen version on fixed predictions. Exact
coverage, missing predictions, verdicts, evidence provenance, and error ownership
belong to the acceptance contract. The current refactor evidence is in
the project's acceptance records.
