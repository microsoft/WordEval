# Shared document access

This package has no dependency on Word Object or Word Revision. It accepts
document paths and opening evidence, never a task instruction or expected
chart/text/table property.

- `package.py` checks ZIP/XML packaging and dispatches structural diagnostics.
- `diagnostics/` groups document, part, chart and relationship rules. The
  extraction preserves the previous rule order and the distinction between
  immediate checks and checks requiring an attributed Word-open failure.
  Passing these checks alone does not establish task success. In particular,
  line-chart `barDir` is diagnostic only after opening failure.
- `word.py` owns Word sessions, normal opening and repair diagnostics. The
  independent probe uses a fresh, byte-identical temporary copy to avoid
  path-specific recovery state. The original is never repaired or saved.
- `open_gate.py` checks reference and prediction independently. An unusable
  reference is an error; confirmed prediction repair/unopenability is a Fail;
  environment failures remain errors.
- `policy.py` assigns prediction/reference/input ownership and normalizes
  evidence. File access errors are not proof of document corruption.

The validated public entry point is `evaluation.Evaluator.evaluate_case()`.
The optional Word Revision CLI uses that same entry point. Scoring adapters
retain their independent semantic rules and may be called directly for
diagnostics, but that bypasses package checks and the external timeout.

`evaluation/isolated.py` bounds native execution and performs scoped cleanup.
Its document-open probe calls this package directly without loading a scoring
backend. Native opening timeouts remain non-verdict errors unless independent
evidence establishes a prediction defect. Word Revision reuses the shared
session and opening primitive; after an opening exception it uses the same
independent gate as Word Object's saved-state acceptance checks.

Structural diagnostics still contain rules originating from historical failures.
Moving a rule into this package does not prove that every matching document is
unusable. Changes to those rules require normal-opening counterexamples and a
separate impact audit; task-semantic checks stay in the scoring method.
