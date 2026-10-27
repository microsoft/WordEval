"""
Evaluation logic for WordEval.

Source of truth for:
- Evaluator: Class-based evaluation orchestrator
- Report: Evaluation report dataclass
- MethodSummary: Summary statistics for a verification method
- load_test_cases: Load test cases from JSON
"""

import json
import os
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, Optional, List, Any

# Import verification methods and base classes
from .document_io.package import validate_docx_output
from .document_io.policy import (
    invalid_output_result, invalid_output_from_typed_word_open_failure,
    invalid_output_after_word_open_failure,
)
from .methods import get_method, list_methods
from .methods.base import VerificationResult, tag_verification_source
from .isolated import run_method_isolated, run_word_open_probe_isolated


SAVED_FALLBACK_OPEN_PROBE_TIMEOUT_SECONDS = 60.0



# =============================================================================
# Data Classes
# =============================================================================

@dataclass
class MethodSummary:
    """Summary statistics for a verification method.

    ``evaluated`` counts every case that produced a verdict or an evaluator
    error (passed + failed + errors), matching the strict SCORE_DEFINITION in
    scripts/evaluate_wordeval.py. Error cases must never silently inflate the
    pass rate by dropping out of the denominator.
    """
    evaluated: int = 0
    passed: int = 0
    failed: int = 0
    missing: int = 0
    errors: int = 0

    @property
    def pass_rate(self) -> float:
        if self.evaluated == 0:
            return 0.0
        return (self.passed / self.evaluated) * 100

    def to_dict(self) -> Dict[str, Any]:
        return {
            'evaluated': self.evaluated,
            'passed': self.passed,
            'failed': self.failed,
            'missing': self.missing,
            'errors': self.errors,
            'pass_rate': round(self.pass_rate, 1),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'MethodSummary':
        return cls(
            evaluated=data.get('evaluated', 0),
            passed=data.get('passed', 0),
            failed=data.get('failed', 0),
            missing=data.get('missing', 0),
            errors=data.get('errors', 0),
        )


@dataclass
class Report:
    """
    Unified evaluation report supporting multiple verification methods.

    Structure:
    - methods: Summary stats per method (from pred.docx)
    - mm_methods: Summary stats per method (from pred_mm.docx, optional)
    - results: Per-case results keyed by "Command/variation"
    """
    test_set: str
    methods: Dict[str, MethodSummary] = field(default_factory=dict)
    mm_methods: Dict[str, MethodSummary] = field(default_factory=dict)
    results: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        d = {
            'test_set': self.test_set,
            'methods': {k: v.to_dict() for k, v in self.methods.items()},
            'results': self.results,
        }
        if self.mm_methods:
            d['mm_methods'] = {k: v.to_dict() for k, v in self.mm_methods.items()}
        if self.metadata:
            d['metadata'] = self.metadata
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Report':
        methods = {}
        for name, method_data in data.get('methods', {}).items():
            methods[name] = MethodSummary.from_dict(method_data)

        mm_methods = {}
        for name, method_data in data.get('mm_methods', {}).items():
            mm_methods[name] = MethodSummary.from_dict(method_data)

        # Handle results - must be dict, not list (old format)
        results = data.get('results', {})
        if isinstance(results, list):
            # Old format - ignore and start fresh
            results = {}

        return cls(
            test_set=data.get('test_set', ''),
            methods=methods,
            mm_methods=mm_methods,
            results=results,
            metadata=data.get('metadata', {}),
        )

    @classmethod
    def load(cls, path: Path) -> Optional['Report']:
        """Load existing report from file."""
        if not path.exists():
            return None
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return cls.from_dict(data)
        except (json.JSONDecodeError, KeyError):
            return None

    def save(self, path: Path):
        """Save report to file atomically.

        Serializes to string in memory first (can't be interrupted mid-JSON),
        then writes to a temp file and replaces the target. If the process is
        killed during the file write, only the temp file is corrupted — the
        target path either keeps its previous valid content or doesn't exist.
        """
        path.parent.mkdir(parents=True, exist_ok=True)
        # Serialize fully in memory — no partial JSON on disk
        data = json.dumps(self.to_dict(), indent=2, sort_keys=True)
        tmp_path = path.with_suffix('.json.tmp')
        with open(tmp_path, 'w', encoding='utf-8') as f:
            f.write(data)
        # Atomic on POSIX; on Windows, replaces target if it exists
        os.replace(tmp_path, path)

    def recalculate_method_summary(self, method_name: str):
        """Recalculate summary stats for a method from results."""
        evaluated = 0
        passed = 0
        failed = 0
        missing = 0
        errors = 0

        for case_data in self.results.values():
            methods_data = case_data.get('methods', {})
            if method_name in methods_data:
                result = methods_data[method_name]
                status = result.get('status', 'unknown')

                if status == 'passed':
                    evaluated += 1
                    passed += 1
                elif status == 'failed':
                    evaluated += 1
                    failed += 1
                elif status == 'error':
                    evaluated += 1
                    errors += 1
                elif status == 'missing':
                    missing += 1

        self.methods[method_name] = MethodSummary(
            evaluated=evaluated,
            passed=passed,
            failed=failed,
            missing=missing,
            errors=errors,
        )

    def recalculate_mm_method_summary(self, method_name: str):
        """Recalculate summary stats for mm_methods from results."""
        evaluated = 0
        passed = 0
        failed = 0
        missing = 0
        errors = 0

        for case_data in self.results.values():
            mm_methods_data = case_data.get('mm_methods', {})
            if method_name in mm_methods_data:
                result = mm_methods_data[method_name]
                status = result.get('status', 'unknown')

                if status == 'passed':
                    evaluated += 1
                    passed += 1
                elif status == 'failed':
                    evaluated += 1
                    failed += 1
                elif status == 'error':
                    evaluated += 1
                    errors += 1
                elif status == 'missing':
                    missing += 1

        if evaluated > 0 or missing > 0:
            self.mm_methods[method_name] = MethodSummary(
                evaluated=evaluated,
                passed=passed,
                failed=failed,
                missing=missing,
                errors=errors,
            )


# =============================================================================
# Evaluator Class
# =============================================================================

class Evaluator:
    """
    Orchestrate verification across test sets.

    Supports multiple parallel verification methods with incremental updates.
    """

    def __init__(
        self,
        methods: List[str] = None,
        case_timeout_seconds: Optional[float] = None,
    ):
        """
        Args:
            methods: List of verification methods to run (default: ["word_object"])
            case_timeout_seconds: Isolate Word-backed methods and enforce this
                per-case timeout. Disabled when None or non-positive.
        """
        method_names = methods or ["word_object"]
        self.methods = [get_method(m) for m in method_names]
        self.case_timeout_seconds = (
            float(case_timeout_seconds)
            if case_timeout_seconds and case_timeout_seconds > 0
            else None
        )

    @staticmethod
    def _saved_fallback_pass_requires_open_gate(
        result: VerificationResult,
    ) -> bool:
        if not result.passed:
            return False
        source = str((result.details or {}).get('verification_source', ''))
        return (
            'saved_ooxml' in source
            and (
                'fallback' in source
                or 'wom_unavailable' in source
            )
            and 'word_com+' not in source
        )

    def _run_saved_fallback_open_probe(
        self,
        pred_path: Path,
        gold_path: Path,
    ) -> Dict[str, Any]:
        timeout = (
            self.case_timeout_seconds
            or SAVED_FALLBACK_OPEN_PROBE_TIMEOUT_SECONDS
        )
        return run_word_open_probe_isolated(
            pred_path,
            gold_path,
            timeout,
        )

    def _enforce_saved_fallback_open_gate(
        self,
        result: VerificationResult,
        pred_path: Path,
        gold_path: Path,
    ) -> VerificationResult:
        """Require explicit full scope and clean-Word evidence for saved passes."""
        if not self._saved_fallback_pass_requires_open_gate(result):
            return result

        original = result.to_dict()
        details = dict(result.details or {})
        source = details.get('verification_source', 'saved_ooxml_fallback')
        if details.get('fallback_contract') != 'complete':
            error = (
                'Saved-state match is not authorized as complete target and '
                'preservation evidence after Word COM is unavailable'
            )
            return VerificationResult(
                passed=False,
                message=f'ERROR: {error}',
                details={
                    **details,
                    'verification_source': source,
                    'saved_ooxml_evidence': original,
                    'missing_evidence': 'complete_fallback_contract',
                },
                error=error,
            )

        probe = self._run_saved_fallback_open_probe(pred_path, gold_path)
        if probe.get('state') == 'completed':
            probe_result = VerificationResult.from_legacy(probe['result'])
            if probe_result.passed:
                result.details = {
                    **details,
                    **dict(probe_result.details or {}),
                    'verification_source': source,
                }
                return result
            probe_result.details = {
                **details,
                'verification_source': source,
                'saved_ooxml_evidence': original,
                **dict(probe_result.details or {}),
            }
            return probe_result

        error = (
            'Saved-state match lacks independent clean-Word normal-open '
            'evidence'
        )
        return VerificationResult(
            passed=False,
            message=f'ERROR: {error}',
            details={
                **details,
                'verification_source': source,
                'saved_ooxml_evidence': original,
                'missing_evidence': 'independent_normal_open',
                'normal_open_probe_state': probe.get('state'),
                'normal_open_probe_error': probe.get('error'),
            },
            error=error,
        )

    def evaluate_case(
        self,
        pred_path: Path,
        gold_path: Path,
        doc_in_path: Path = None,
        spec: Dict = None
    ) -> Dict[str, Dict[str, Any]]:
        """
        Evaluate a single test case with all configured methods.

        Returns:
            Dict mapping method name to result dict with status and details
        """
        results = {}
        try:
            invalid_output = validate_docx_output(pred_path)
            if not invalid_output:
                reference_issue = validate_docx_output(gold_path)
                if reference_issue:
                    return {
                        method.name: {
                            'status': 'error', 'passed': False,
                            'message': 'ERROR: Reference document failed package validation',
                            'details': {'verification_source': 'package_preflight',
                                        'open_failure_owner': 'gold',
                                        'reference_issue': reference_issue},
                            'error': 'Invalid benchmark reference document',
                        }
                        for method in self.methods if method.supports_spec(spec)
                    }
        except Exception as exc:
            error = f'Package preflight failed: {type(exc).__name__}: {exc}'
            for method in self.methods:
                if method.supports_spec(spec):
                    results[method.name] = {
                        'status': 'error',
                        'passed': False,
                        'message': f'ERROR: {error}',
                        'details': {
                            'verification_source': 'package_preflight',
                            'error_owner': 'evaluator',
                            'exception_type': type(exc).__name__,
                        },
                        'error': error,
                    }
            return results

        for method in self.methods:
            if method.supports_spec(spec):
                if invalid_output:
                    results[method.name] = invalid_output_result(invalid_output)
                    continue

                try:
                    if self.case_timeout_seconds and method.requires_word:
                        isolated = run_method_isolated(
                            method.name,
                            pred_path,
                            gold_path,
                            doc_in_path,
                            spec,
                            self.case_timeout_seconds,
                        )
                        if isolated['state'] == 'completed':
                            result = VerificationResult.from_legacy(
                                isolated['result']
                            )
                            if isolated.get('worker_cleanup_timed_out'):
                                result.details = dict(result.details or {})
                                result.details['worker_cleanup_timed_out'] = True
                            if isolated.get('cleanup_error'):
                                result.details = dict(result.details or {})
                                result.details['cleanup_error'] = isolated[
                                    'cleanup_error'
                                ]
                        elif isolated['state'] == 'timeout':
                            fallback = getattr(
                                method, 'verify_saved_fallback', None
                            )
                            fallback_result = (
                                fallback(
                                    pred_path, gold_path, doc_in_path, spec
                                )
                                if fallback else None
                            )
                            if fallback_result is None:
                                results[method.name] = {
                                    'status': 'error',
                                    'passed': False,
                                    'message': (
                                        'ERROR: Word COM timed out after '
                                        f'{self.case_timeout_seconds:g}s and no '
                                        'saved-state fallback is registered'
                                    ),
                                    'details': {
                                        'verification_source': 'word_com_timeout',
                                        'timeout_seconds': self.case_timeout_seconds,
                                        **(
                                            {'cleanup_error': isolated['cleanup_error']}
                                            if isolated.get('cleanup_error') else {}
                                        ),
                                    },
                                    'error': isolated.get('error') or 'Word COM timeout',
                                }
                                continue
                            result = tag_verification_source(
                                fallback_result,
                                'saved_ooxml_timeout_fallback',
                                word_com_timeout_seconds=self.case_timeout_seconds,
                                word_com_error=(
                                    isolated.get('error') or 'Word COM timeout'
                                ),
                                **(
                                    {'cleanup_error': isolated['cleanup_error']}
                                    if isolated.get('cleanup_error') else {}
                                ),
                            )
                        else:
                            results[method.name] = {
                                'status': 'error',
                                'passed': False,
                                'message': 'ERROR: Isolated evaluator worker failed',
                                'details': {
                                    'verification_source': 'isolated_worker_error',
                                    **(
                                        {'cleanup_error': isolated['cleanup_error']}
                                        if isolated.get('cleanup_error') else {}
                                    ),
                                },
                                'error': isolated.get('error', 'worker error'),
                            }
                            continue
                    else:
                        result = method.verify(
                            pred_path, gold_path, doc_in_path, spec
                        )
                    result = self._enforce_saved_fallback_open_gate(
                        result,
                        pred_path,
                        gold_path,
                    )
                    # Flatten: merge result fields with status
                    result_dict = result.to_dict()
                    if result.error:
                        result_dict['status'] = 'error'
                    else:
                        result_dict['status'] = (
                            'passed' if result.passed else 'failed'
                        )
                    open_failure_owner = (result_dict.get('details') or {}).get(
                        'open_failure_owner'
                    )
                    open_failure_invalid = invalid_output_from_typed_word_open_failure(
                        result_dict
                    )
                    if (
                        open_failure_invalid is None
                        and open_failure_owner in {None, 'prediction'}
                    ):
                        open_failure_invalid = invalid_output_after_word_open_failure(
                            pred_path,
                            result_dict.get('error'),
                        )
                    if open_failure_invalid is not None:
                        result_dict = open_failure_invalid
                    results[method.name] = result_dict
                except Exception as e:
                    open_failure_invalid = invalid_output_after_word_open_failure(
                        pred_path,
                        str(e),
                    )
                    if open_failure_invalid is not None:
                        results[method.name] = open_failure_invalid
                    else:
                        results[method.name] = {
                            'status': 'error',
                            'passed': False,
                            'message': f'ERROR: {e}',
                            'details': {},
                            'error': str(e),
                        }

        return results

    def evaluate_test_set(
        self,
        test_set_dir: Path,
        results_dir: Path,
        report_path: Path = None,
        commands: List[str] = None
    ) -> Report:
        """
        Run evaluation and merge results into existing report.

        Args:
            test_set_dir: Path to test set (contains test_cases.json)
            results_dir: Path to model results (contains pred.docx files)
            report_path: Path to existing report (will be loaded and updated)
            commands: Optional list of specific commands to evaluate

        Returns:
            Updated Report with results from all methods
        """
        test_cases = load_test_cases(test_set_dir)
        if not test_cases:
            raise ValueError(f"test_cases.json not found in {test_set_dir}")

        test_set_name = test_cases.get("test_set", test_set_dir.name)
        commands_data = test_cases.get("commands", {})

        # Load existing report or create new one
        report = None
        if report_path:
            report = Report.load(report_path)
        if not report:
            report = Report(
                test_set=test_set_name,
                methods={},
                results={},
            )

        for idMso, variations in commands_data.items():
            if commands and idMso not in commands:
                continue

            for variation, data in variations.items():
                case_id = f"{idMso}/{variation}"
                instruction = data.get("instruction", "")
                doc_in_rel = data.get("doc_in", "")
                gold_rel = data.get("gold", "")
                spec = data.get("spec", {})

                spec['idMso'] = idMso
                spec['variation'] = variation

                # Resolve paths
                gold_path = test_set_dir / gold_rel
                pred_path = results_dir / idMso / variation / "pred.docx"
                doc_in_path = test_set_dir / doc_in_rel if doc_in_rel else None

                # Initialize case entry if not exists, always update instruction from test_cases.json
                if case_id not in report.results:
                    report.results[case_id] = {
                        'instruction': instruction,
                        'methods': {},
                    }
                else:
                    # Always update instruction from test_cases.json (may have changed)
                    report.results[case_id]['instruction'] = instruction

                # Check if pred exists
                if not pred_path.exists():
                    for method in self.methods:
                        report.results[case_id]['methods'][method.name] = {
                            'status': 'missing',
                        }
                else:
                    # Run evaluation with all methods
                    method_results = self.evaluate_case(
                        pred_path=pred_path,
                        gold_path=gold_path,
                        doc_in_path=doc_in_path if doc_in_path and doc_in_path.exists() else None,
                        spec=spec
                    )
                    # Merge results
                    report.results[case_id]['methods'].update(method_results)

                # Check if pred_mm.docx exists (multimodal prediction)
                pred_mm_path = results_dir / idMso / variation / "pred_mm.docx"
                if pred_mm_path.exists():
                    if 'mm_methods' not in report.results[case_id]:
                        report.results[case_id]['mm_methods'] = {}
                    mm_method_results = self.evaluate_case(
                        pred_path=pred_mm_path,
                        gold_path=gold_path,
                        doc_in_path=doc_in_path if doc_in_path and doc_in_path.exists() else None,
                        spec=spec
                    )
                    report.results[case_id]['mm_methods'].update(mm_method_results)

        # Recalculate summaries for methods we ran
        for method in self.methods:
            report.recalculate_method_summary(method.name)
            report.recalculate_mm_method_summary(method.name)

        return report


# =============================================================================
# Utility Functions
# =============================================================================

def load_test_cases(test_set_dir: Path) -> dict:
    """Load test cases from test_cases.json."""
    test_cases_path = test_set_dir / "test_cases.json"
    if not test_cases_path.exists():
        return None

    with open(test_cases_path, "r", encoding="utf-8") as f:
        return json.load(f)


