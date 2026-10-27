"""Own Word automation processes and document lifetime independently of readers."""

import os
import hashlib
import tempfile
from pathlib import Path
import time
from typing import Any, Dict
from .errors import DocumentOpenError, NormalOpenError


def open_normal_document(word, path, **options):
    """Use the submitted file; one COM exception is not a corruption verdict."""
    options.update(OpenAndRepair=False, AddToRecentFiles=False, ConfirmConversions=False)
    try:
        return word.Documents.Open(os.path.abspath(path), **options)
    except Exception as exc:
        raise NormalOpenError(path, exc) from exc


class WordSession:
    """Sequential, ownership-aware Word automation session."""

    _WRITE_PASSWORD_PROBE = "WordEval-no-write-access"

    def __init__(self, max_retries: int = 3, retry_delay: float = 1.0):
        """
        Args:
            max_retries: Maximum retry attempts for COM operations
            retry_delay: Delay between retries in seconds
        """
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self._word = None
        self._baseline_word_process_ids = self._running_word_process_ids()
        self._owned_word_process_ids = set()

    @staticmethod
    def _running_word_process_ids() -> set[int]:
        """Return live automation-owned WINWORD PIDs.

        Only COM automation instances (``/automation`` or ``-Embedding`` on
        the command line) are ownership/cleanup candidates. An interactive
        Word session the user opens mid-run must never be claimed, waited on,
        or terminated by the evaluator.
        """
        try:
            import psutil
        except ImportError:
            return set()
        result = set()
        # Prefetch only the name (one system snapshot); read the command line
        # per WINWORD candidate under its own handler.
        for process in psutil.process_iter(["name"]):
            try:
                if (process.info.get("name") or "").lower() != "winword.exe":
                    continue
                command_line = " ".join(process.cmdline() or []).casefold()
                if "/automation" not in command_line and "-embedding" not in command_line:
                    continue
                result.add(int(process.pid))
            except psutil.Error:
                continue
        return result

    def _get_word(self):
        """Get or create a healthy Word application instance."""
        if self._word is not None and not self.is_word_responsive():
            # A crashed or wedged instance must not serve further COM calls:
            # on a dead proxy every document open fails for file-independent
            # reasons, which downstream code can misclassify as document
            # problems (e.g. "repair required" on a perfectly valid file).
            self._cleanup_word()
        if self._word is None:
            if not self._cleanup_word():
                raise RuntimeError("Prior evaluator-owned Word process did not exit")
            import win32com.client

            before = self._running_word_process_ids()
            # Evaluation owns this Word process. Dispatch can attach to an
            # unrelated interactive instance and also makes a failed/corrupt
            # document open contaminate later repair attempts.
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = 0  # Suppress alerts
            self._word.AutomationSecurity = 3
            after = self._running_word_process_ids()
            self._owned_word_process_ids.update(after - before)
            self._owned_word_process_ids.update(after - self._baseline_word_process_ids)
        return self._word

    def _cleanup_word(self, wait_timeout_seconds: float = 60.0) -> bool:
        """Quit the owned Word instance and wait for its process to exit.

        ``Application.Quit`` can return while one or more out-of-process Word
        servers created by this evaluator are still alive.  A refresh
        observation must not start another instance during that interval, so
        wait for the exact window PID plus every WINWORD PID that appeared
        after this core's baseline.  Baseline user processes are never owned,
        waited on, or terminated.  No fixed sleep is used.
        """
        owned_pid = None
        quit_succeeded = True
        if self._word is not None:
            try:
                import win32process

                owned_pid = int(win32process.GetWindowThreadProcessId(int(self._word.Hwnd))[1])
            except Exception:
                owned_pid = None
            if owned_pid:
                self._owned_word_process_ids.add(owned_pid)

            try:
                self._word.Quit()
            except Exception:
                quit_succeeded = False
            finally:
                self._word = None

        try:
            import psutil
        except ImportError:
            return quit_succeeded and owned_pid is None

        deadline = time.monotonic() + wait_timeout_seconds
        while True:
            candidate_pids = self._owned_word_process_ids | (
                self._running_word_process_ids() - self._baseline_word_process_ids
            )
            processes = []
            for process_id in candidate_pids:
                try:
                    processes.append(psutil.Process(process_id))
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            if not processes:
                self._owned_word_process_ids.clear()
                return quit_succeeded

            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return False
            _, alive = psutil.wait_procs(processes, timeout=remaining)
            if alive:
                return False
            self._owned_word_process_ids.difference_update(candidate_pids)

    def restart_word(self):
        """Restart Word application (cleanup and recreate)."""
        self._cleanup_word()
        return self._get_word()

    def is_word_responsive(self) -> bool:
        """Check if Word is still responsive."""
        if self._word is None:
            return False
        try:
            _ = self._word.Version
            return True
        except:
            return False

    def _open_document_with_status(
        self, docx_path: str, read_only: bool = True, allow_repair: bool = True
    ):
        """
        Open a document with retry logic and return repair status.

        Args:
            docx_path: Path to the document
            read_only: Open in read-only mode
            allow_repair: If True, try OpenAndRepair when normal open fails

        Returns:
            Tuple of (document, needed_repair) where needed_repair is True if
            the document could only be opened with repair
        """
        word = self._get_word()
        abs_path = os.path.abspath(docx_path)
        if os.path.isabs(docx_path):
            try:
                rel_path = os.path.relpath(docx_path)
            except ValueError:
                # Windows: document on a different drive than the working
                # directory has no relative form; keep the original path.
                rel_path = docx_path
        else:
            rel_path = docx_path

        open_kwargs = {"ReadOnly": read_only}
        if not read_only:
            open_kwargs["WritePasswordDocument"] = self._WRITE_PASSWORD_PROBE

        # First try opening without repair on the current instance.
        last_normal_error = None
        for attempt in range(self.max_retries):
            try:
                doc = open_normal_document(word, abs_path, **open_kwargs)
                return doc, False  # Opened successfully without repair
            except Exception as e:
                last_normal_error = e
                if attempt < self.max_retries - 1:
                    time.sleep(self.retry_delay)

        # Every normal open on the shared instance failed. That instance may
        # itself be dead or wedged (e.g. Word crashed earlier in the case),
        # making the failures file-independent. "Needs repair" is a
        # verdict-affecting classification, so it must be established against
        # the file, not against a broken application: retry a NORMAL open in a
        # fresh process before attributing anything to the document. A restart
        # failure raises RuntimeError (evaluator-owned), deliberately not
        # DocumentOpenError.
        word = self.restart_word()
        try:
            doc = open_normal_document(word, abs_path, **open_kwargs)
            return doc, False
        except Exception as fresh_normal_error:
            last_normal_error = fresh_normal_error

        if not allow_repair:
            raise DocumentOpenError(
                abs_path,
                f"Failed to open {rel_path}: {last_normal_error}",
                normal_error=last_normal_error,
            )

        # A failed normal open can leave Word's document loader in a state
        # where OpenAndRepair fails or hangs on the same application. Repair
        # must run in a process with no failed open on it; this is also how
        # Word's manual Open and Repair path behaves reliably in practice.
        # Restart outside the try: a restart failure is evaluator-owned.
        word = self.restart_word()
        try:
            doc = word.Documents.Open(abs_path, OpenAndRepair=True, **open_kwargs)
            return doc, True  # Opened with repair
        except Exception as repair_e:
            message = (
                f"Failed to open {rel_path} normally ({last_normal_error}) "
                "or with repair in a fresh Word instance: "
                f"{repair_e}"
            )
            raise DocumentOpenError(
                abs_path,
                message,
                normal_error=last_normal_error,
                repair_error=repair_e,
            )

    def _open_document(self, docx_path: str, read_only: bool = True):
        """Open a document normally; repaired output is not valid evaluation input.

        Repair mode remains available through ``_open_document_with_status`` for
        diagnostics. Formal Word Object Model extraction must not silently score
        a repaired in-memory document as if the submitted file opened normally.
        """
        doc, needed_repair = self._open_document_with_status(
            docx_path, read_only, allow_repair=True
        )
        if needed_repair:
            self._close_document(doc)
            abs_path = os.path.abspath(docx_path)
            raise DocumentOpenError(
                abs_path,
                f"Document requires Word repair: {abs_path}",
                repair_succeeded=True,
            )
        return doc

    @classmethod
    def probe_clean_openability(
        cls,
        docx_path: str,
        *,
        max_retries: int = 2,
        retry_delay: float = 0.25,
    ) -> Dict[str, Any]:
        """Probe normal/repair open in a fresh automation-owned Word process.

        This is deliberately a Word observation, not a strict-schema oracle.
        The caller assigns prediction-versus-gold ownership.
        """
        probe = cls(max_retries=max_retries, retry_delay=retry_delay)
        doc = None
        staging = None
        source = Path(docx_path).resolve()
        result = {}
        try:
            data = source.read_bytes()
            digest = hashlib.sha256(data).hexdigest()
            # Word can remember a crash against a path. Diagnose identical bytes
            # at a new path instead of deleting the user's global recovery state.
            staging = tempfile.TemporaryDirectory(prefix="wordeval-open-probe-")
            staged = Path(staging.name) / source.name
            staged.write_bytes(data)
            doc, needed_repair = probe._open_document_with_status(
                str(staged), read_only=True, allow_repair=True,
            )
            result = {
                "path": str(source), "sha256": digest,
                "state": "repair_required" if needed_repair else "normal_opened",
                "normal_opened": not needed_repair, "repair_opened": needed_repair,
                "fresh_path": True,
            }
            if hashlib.sha256(source.read_bytes()).hexdigest() != digest:
                raise RuntimeError("Document bytes changed during openability probe")
        except DocumentOpenError as exc:
            result = {**exc.to_details(), "path": str(source),
                      "sha256": digest, "fresh_path": True}
        except Exception as exc:
            result = {"path": str(source), "state": "environment_error",
                      "normal_opened": False, "repair_opened": False,
                      "environment_error": str(exc)}
        finally:
            probe._close_document(doc)
            cleaned = probe._cleanup_word()
            try:
                if staging is not None:
                    staging.cleanup()
                    if hashlib.sha256(source.read_bytes()).hexdigest() != digest:
                        raise RuntimeError('Document bytes changed during openability probe')
                if not cleaned:
                    raise RuntimeError("Owned Word process did not exit after openability probe")
            except Exception as exc:
                result = {"path": str(source), "state": "environment_error",
                          "normal_opened": False, "repair_opened": False,
                          "environment_error": str(exc), "observation": result}
        return result

    def check_file_needs_repair(self, docx_path: str) -> bool:
        """
        Check if a file requires repair to open.

        Returns True if the file cannot be opened normally and needs repair.
        Returns False if the file can be opened without repair.
        Raises RuntimeError if the file cannot be opened even with repair.
        """
        doc = None
        try:
            doc, needed_repair = self._open_document_with_status(
                docx_path, read_only=True, allow_repair=True
            )
            return needed_repair
        finally:
            self._close_document(doc)

    def _close_document(self, doc):
        """Close a document safely."""
        if doc is not None:
            try:
                doc.Close(False)  # False = don't save changes
            except:
                pass

    def __del__(self):
        """Cleanup on destruction."""
        self._cleanup_word()
