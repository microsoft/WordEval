"""Case-level process isolation for Word COM verification."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Optional

try:
    import psutil
except ImportError:  # Checked fail-closed in run_method_isolated.
    psutil = None

# The worker must import THIS evaluation package, not whatever `evaluation`
# happens to resolve from the operator's working directory (a frozen snapshot
# silently running the live repo's verifiers defeats snapshot fingerprints).
_PACKAGE_PARENT = Path(__file__).resolve().parent.parent


def _automation_word_processes() -> dict[int, float]:
    if psutil is None:
        return {}
    processes = {}
    # Do not request attributes through process_iter().  On Windows, psutil
    # populates ``candidate.info`` before the loop body runs; an unrelated
    # protected process can therefore raise AccessDenied outside our per-item
    # exception handler and abort Word cleanup entirely.
    for candidate in psutil.process_iter():
        try:
            if str(candidate.name() or "").upper() != "WINWORD.EXE":
                continue
            command_line = " ".join(candidate.cmdline() or []).casefold()
            if "/automation" not in command_line and "-embedding" not in command_line:
                continue
            processes[candidate.pid] = float(candidate.create_time() or 0)
        except psutil.Error:
            pass
    return processes


def _terminate_process_tree(process: subprocess.Popen[str]) -> None:
    if psutil is None:
        process.kill()
        return
    try:
        parent = psutil.Process(process.pid)
        children = parent.children(recursive=True)
    except psutil.Error:
        children = []
    for child in reversed(children):
        try:
            child.kill()
        except psutil.Error:
            pass
    try:
        process.kill()
    except OSError:
        pass


def _cleanup_new_word_processes(before: set[int], started: float) -> None:
    deadline = time.monotonic() + 3.0
    while time.monotonic() < deadline:
        killed = []
        for pid, created in _automation_word_processes().items():
            if pid in before or created < started - 1:
                continue
            try:
                candidate = psutil.Process(pid)
                candidate.kill()
                killed.append(candidate)
            except psutil.Error:
                pass
        if not killed:
            return
        psutil.wait_procs(killed, timeout=1)
        time.sleep(0.1)


def _last_json_line(output: str) -> Optional[dict[str, Any]]:
    for line in reversed(output.splitlines()):
        if not line.strip():
            continue
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    return None


def run_method_isolated(
    method: str,
    pred_path: Path,
    gold_path: Path,
    doc_in_path: Optional[Path],
    spec: Optional[dict[str, Any]],
    timeout_seconds: float,
) -> dict[str, Any]:
    """Run one method with a hard timeout and scoped Word cleanup."""
    if psutil is None:
        # Without psutil, a timed-out DCOM Word instance cannot be found or
        # killed (it is not a child of the worker) and leaks one WINWORD.EXE
        # per timeout. Refuse to run rather than degrade silently.
        raise RuntimeError(
            "psutil is required for isolated Word evaluation: timed-out "
            "Word processes cannot be cleaned up without it. "
            "Install it with: pip install psutil"
        )
    action = (
        "probe_word_openability"
        if method == "__word_open_probe__"
        else "verify"
    )
    payload = json.dumps({
        "action": action,
        "method": None if action == "probe_word_openability" else method,
        "pred_path": str(pred_path.resolve()),
        "gold_path": str(gold_path.resolve()),
        "doc_in_path": str(doc_in_path.resolve()) if doc_in_path else None,
        "spec": spec,
        "evaluation_root": str(_PACKAGE_PARENT / "evaluation"),
    }, ensure_ascii=False)
    before_word = set(_automation_word_processes())
    started = time.time()
    worker_env = os.environ.copy()
    existing_pythonpath = worker_env.get("PYTHONPATH")
    worker_env["PYTHONPATH"] = (
        str(_PACKAGE_PARENT)
        if not existing_pythonpath
        else str(_PACKAGE_PARENT) + os.pathsep + existing_pythonpath
    )
    process = subprocess.Popen(
        [sys.executable, "-u", "-m", "evaluation.isolated_worker"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(_PACKAGE_PARENT),
        env=worker_env,
    )
    cleanup_errors: list[str] = []

    def best_effort_cleanup(action, label: str) -> None:
        try:
            action()
        except Exception as exc:
            cleanup_errors.append(f"{label}: {type(exc).__name__}: {exc}")

    def with_cleanup_diagnostics(result: dict[str, Any]) -> dict[str, Any]:
        if cleanup_errors:
            result["cleanup_error"] = "; ".join(cleanup_errors)
        return result

    try:
        stdout, stderr = process.communicate(
            input=payload, timeout=timeout_seconds
        )
    except subprocess.TimeoutExpired as exc:
        best_effort_cleanup(
            lambda: _terminate_process_tree(process),
            "terminate timed-out worker",
        )
        # A Word instance started through DCOM is not necessarily a child of
        # the Python worker.  It can also keep inherited pipe handles alive,
        # causing an unbounded second communicate() after the worker timeout.
        # Close the scoped automation processes before draining the pipes and
        # keep that drain bounded as a second line of defense.
        best_effort_cleanup(
            lambda: _cleanup_new_word_processes(before_word, started),
            "cleanup timed-out Word processes",
        )
        try:
            stdout_after, stderr_after = process.communicate(timeout=5)
        except subprocess.TimeoutExpired:
            best_effort_cleanup(
                lambda: _terminate_process_tree(process),
                "terminate worker after pipe timeout",
            )
            best_effort_cleanup(
                lambda: _cleanup_new_word_processes(before_word, started),
                "cleanup Word processes after pipe timeout",
            )
            try:
                stdout_after, stderr_after = process.communicate(timeout=2)
            except subprocess.TimeoutExpired:
                stdout_after, stderr_after = "", "worker pipe cleanup timed out"
        captured = exc.stdout or ""
        if isinstance(captured, bytes):
            captured = captured.decode(errors="replace")
        if stdout_after and stdout_after not in captured:
            captured += stdout_after
        completed_result = _last_json_line(captured)
        if completed_result is not None:
            return with_cleanup_diagnostics({
                "state": "completed",
                "result": completed_result,
                "worker_cleanup_timed_out": True,
            })
        return with_cleanup_diagnostics({
            "state": "timeout",
            "error": (stderr_after or "Word COM case timeout").strip(),
        })

    best_effort_cleanup(
        lambda: _cleanup_new_word_processes(before_word, started),
        "cleanup completed Word processes",
    )
    result = _last_json_line(stdout)
    if process.returncode == 0 and result is not None:
        return with_cleanup_diagnostics({"state": "completed", "result": result})
    return with_cleanup_diagnostics({
        "state": "worker_error",
        "error": (stderr or stdout or "isolated worker returned no result").strip(),
        "returncode": process.returncode,
    })


def run_word_open_probe_isolated(
    pred_path: Path,
    gold_path: Path,
    timeout_seconds: float,
) -> dict[str, Any]:
    """Probe pred/gold openability in a disposable clean Word worker."""
    return run_method_isolated(
        "__word_open_probe__",
        pred_path,
        gold_path,
        None,
        None,
        timeout_seconds,
    )
