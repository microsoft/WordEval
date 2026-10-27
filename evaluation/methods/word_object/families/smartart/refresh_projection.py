"""families / smartart / refresh_projection with explicit evidence dependencies."""

from evaluation.methods.word_object.families.smartart import (
    saved_projection as families_smartart_saved_projection,
)
from typing import Dict, Any
import hashlib
import os
import shutil
import tempfile


def extract_smartart_visible_projection(docx_path: str, *, create_session) -> Dict[str, Any]:
    """Build a Word-refreshed SmartArt projection on an isolated temp copy.

    Structural SmartArt edits can leave ``drawing*.xml`` in a stale raw
    order even though a normal Word open exposes and renders the requested
    hierarchy.  Conversely, the live ``SmartArt.AllNodes`` collection can
    contain a node that Word does not render.  To distinguish those cases,
    copy the submitted package, normal-open it in a dedicated Word process,
    SaveAs to a second temporary DOCX, then read the refreshed visible node
    text/order/styles plus semantic gallery IDs from that saved package.

    The source package is never opened for write and is hashed before and
    after the operation.  Missing or indeterminate evidence is returned as
    an attributable extraction error; callers must not infer a pass.
    """

    source = os.path.abspath(docx_path)

    def file_sha256(path: str) -> str:
        digest = hashlib.sha256()
        with open(path, "rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    result: Dict[str, Any] = {
        "found": False,
        "refresh_method": "isolated_temp_copy_word_saveas",
    }
    refresh_core = create_session(max_retries=1, retry_delay=0.0)
    word = None
    doc = None
    source_hash_before = None
    try:
        source_hash_before = file_sha256(source)
        result["source_sha256_before"] = source_hash_before
        with tempfile.TemporaryDirectory(
            prefix="wordeval_smartart_refresh_",
            ignore_cleanup_errors=True,
        ) as temp_dir:
            temp_copy = os.path.join(temp_dir, "submitted_copy.docx")
            refreshed_copy = os.path.join(temp_dir, "refreshed_copy.docx")
            shutil.copy2(source, temp_copy)
            result["temporary_copy_sha256_before"] = file_sha256(temp_copy)

            # Do not use _open_document here: this projection must never
            # enter OpenAndRepair.  Formal callers have already completed
            # the package/normal-open gate, and this fresh process is an
            # independent refresh observation.
            word = refresh_core._get_word()
            doc = word.Documents.Open(
                temp_copy,
                ReadOnly=False,
                OpenAndRepair=False,
                AddToRecentFiles=False,
            )
            try:
                result["page_count"] = int(doc.ComputeStatistics(2))
            except Exception:
                result["page_count"] = None
            doc.SaveAs2(
                refreshed_copy,
                FileFormat=16,
                AddToRecentFiles=False,
            )
            doc.Close(False)
            doc = None
            # Release the local Application proxy before cleanup. Keeping
            # it alive while waiting for WINWORD to exit can extend the
            # lifetime of the out-of-process COM server after Quit().
            word = None

            drawing = families_smartart_saved_projection.extract_smartart_drawing_info(
                refreshed_copy
            )
            gallery = families_smartart_saved_projection.extract_smartart_gallery_identities(
                refreshed_copy
            )
            result["temporary_refreshed_sha256"] = file_sha256(refreshed_copy)
            result["drawing_found"] = bool(drawing.get("found"))
            result["visible_node_order"] = list(drawing.get("node_order") or [])
            result["visible_node_count"] = len(result["visible_node_order"])
            result["visible_node_styles"] = list(drawing.get("node_styles") or [])
            result["gallery_identities_found"] = bool(gallery.get("found"))
            result["gallery_identities"] = list(gallery.get("identities") or [])
            result["found"] = bool(drawing.get("found") and result["visible_node_order"])
            if drawing.get("error"):
                result["error"] = drawing["error"]
            elif gallery.get("error"):
                result["error"] = gallery["error"]
            elif not result["found"]:
                result["error"] = "Refreshed SmartArt drawing projection is empty"
    except Exception as exc:
        result["error"] = str(exc)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
            finally:
                doc = None
        word = None
        refresh_cleanup_complete = refresh_core._cleanup_word()
        result["refresh_process_exited"] = refresh_cleanup_complete
        if not refresh_cleanup_complete:
            result["found"] = False
            result["error"] = "Dedicated SmartArt refresh Word process did not exit"
        try:
            source_hash_after = file_sha256(source)
            result["source_sha256_after"] = source_hash_after
            result["source_unchanged"] = bool(
                source_hash_before and source_hash_before == source_hash_after
            )
            if source_hash_before and not result["source_unchanged"]:
                result["found"] = False
                result["error"] = "Canonical SmartArt source hash changed during refresh"
        except Exception as exc:
            result["found"] = False
            result["source_unchanged"] = False
            result["error"] = "Could not verify canonical SmartArt source hash: " f"{exc}"
    return result
