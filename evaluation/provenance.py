"""Portable evaluation identities, recorded without starting an Office process."""
from __future__ import annotations

import hashlib
import importlib.metadata
import json
import platform
import re
from pathlib import Path
import sys

SNAPSHOT_FINGERPRINT_ALGORITHM = (
    "sha256 over sorted relative-path/NUL/file-sha256 rows; UTF-8 text line endings are normalized to LF, "
    "and generated_at JSON fields are normalized only for exported wordeval and wordeval_inputs trees"
)
_TEXT_SUFFIXES = {".csv", ".json", ".jsonl", ".md", ".py", ".toml", ".txt", ".yaml", ".yml"}


def _normalize_generated_at(value):
    if isinstance(value, dict):
        return {key: "<normalized-generated-at>" if key == "generated_at" else _normalize_generated_at(item)
                for key, item in value.items()}
    if isinstance(value, list):
        return [_normalize_generated_at(item) for item in value]
    return value


def snapshot_artifact_hash(path: Path, *, normalize_json=False) -> str:
    """The existing frozen-release hash format, distinct from raw-byte identity."""
    data = path.read_bytes()
    if normalize_json and path.suffix.lower() == ".json":
        data = json.dumps(_normalize_generated_at(json.loads(data.decode("utf-8"))),
                          ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    elif path.suffix.lower() in _TEXT_SUFFIXES:
        data = data.decode("utf-8").replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")
    return hashlib.sha256(data).hexdigest()


def snapshot_tree_fingerprint(root: Path, *, normalize_json=False) -> str:
    if not root.is_dir():
        raise ValueError(f"Release tree is missing: {root}")
    rows = []
    for path in sorted(root.rglob("*"), key=lambda p: (p.relative_to(root).as_posix().casefold(),
                                                     p.relative_to(root).as_posix())):
        if not path.is_file() or "__pycache__" in path.parts or path.suffix == ".pyc":
            continue
        relative = path.relative_to(root).as_posix()
        rows.append(f"{relative}\0{snapshot_artifact_hash(path, normalize_json=normalize_json)}\n".encode("utf-8"))
    return hashlib.sha256(b"".join(rows)).hexdigest()


def validate_release_manifest(manifest_path, *, runtime_root, runner_path, wordeval_dir,
                              selected_tracks, read_oracle, predictions_dir=None):
    """Validate the scoring artifacts and complete oracle before any Word work.

    This checks integrity against an explicitly supplied manifest. It does not
    certify Office compatibility, authorise publication or verify predictions.
    """
    manifest_path = Path(manifest_path).resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict) or manifest.get("fingerprint_algorithm") != SNAPSHOT_FINGERPRINT_ALGORITHM:
        raise ValueError("Unsupported release manifest fingerprint algorithm")
    if manifest.get("schema") != "wordeval" or manifest.get("mutable") is not False:
        raise ValueError("Release manifest must declare immutable WordEval assets")
    if not isinstance(manifest.get("id"), str) or not manifest["id"]:
        raise ValueError("Release manifest id is missing")
    artifacts, expected = manifest.get("artifacts", {}), manifest.get("fingerprints", {})
    if not isinstance(artifacts, dict) or not isinstance(expected, dict):
        raise ValueError("Malformed release artifact declarations")
    checked = {}
    for key, fingerprint_key, actual_path in (
        ("evaluation", "evaluation", runtime_root),
        ("evaluate_wordeval", "evaluate_wordeval.py", runner_path),
        ("wordeval", "wordeval", wordeval_dir),
    ):
        declared = artifacts.get(key)
        if not isinstance(declared, str) or not declared or Path(declared).is_absolute():
            raise ValueError(f"Invalid release artifact path: {key}")
        path = (manifest_path.parent / declared).resolve()
        if not path.is_relative_to(manifest_path.parent) or path != Path(actual_path).resolve():
            raise ValueError(f"Actual {key} does not match the release manifest path")
        wanted = expected.get(fingerprint_key)
        if not isinstance(wanted, str) or not re.fullmatch(r"[0-9a-f]{64}", wanted):
            raise ValueError(f"Missing release fingerprint: {fingerprint_key}")
        actual = (snapshot_artifact_hash(path) if key == "evaluate_wordeval" else
                  snapshot_tree_fingerprint(path, normalize_json=(key == "wordeval")))
        if actual != wanted:
            raise ValueError(f"Release fingerprint mismatch: {fingerprint_key}")
        checked[fingerprint_key] = actual
    counts = manifest.get("case_counts")
    if not isinstance(counts, dict) or "total" not in counts:
        raise ValueError("Release manifest case_counts are missing")
    tracks = {name: count for name, count in counts.items() if name != "total"}
    if not tracks or any(not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]*", name) or type(count) is not int or count <= 0
                         for name, count in tracks.items()):
        raise ValueError("Invalid release track counts")
    if type(counts["total"]) is not int or counts["total"] != sum(tracks.values()):
        raise ValueError("Release total case count does not match track counts")
    discovered = {p.parent.name for p in Path(wordeval_dir).glob("*/oracle.jsonl")}
    if discovered != set(tracks) or not set(selected_tracks) <= set(tracks):
        raise ValueError("Release track set does not match manifest")
    expected_slots = set()
    for track, count in tracks.items():
        folder = Path(wordeval_dir) / track
        rows = read_oracle(folder / "oracle.jsonl")
        if len(rows) != count:
            raise ValueError(f"Release case count mismatch: {track}")
        for row in rows:
            expected_slots.add((track, row["case_id"]))
            for field in ("instruction", "source_command", "source_variant", "input_document", "reference_document"):
                if not isinstance(row.get(field), str) or not row[field].strip():
                    raise ValueError(f"Missing required oracle field {field}: {track}/{row['case_id']}")
            if not isinstance(row.get("evaluation_spec"), dict) or not row["evaluation_spec"]:
                raise ValueError(f"Missing evaluation_spec: {track}/{row['case_id']}")
            for field in ("input_document", "reference_document"):
                if not (folder / row[field]).is_file():
                    raise ValueError(f"Missing release document: {track}/{row['case_id']}/{field}")
    if predictions_dir is not None:
        observed_slots = {(path.parent.parent.name, path.parent.name)
                          for path in Path(predictions_dir).glob("*/*/pred.docx")}
        unknown = sorted(observed_slots - expected_slots)
        if unknown:
            raise ValueError(f"Unknown prediction case identities: {unknown[:5]}")
    return {"verified": True, "release_id": manifest["id"], "manifest": file_identity(manifest_path),
            "checked_artifacts": checked, "case_counts": counts,
            "complete_track_selection": set(selected_tracks) == set(tracks)}


def file_identity(path: Path) -> dict:
    path = Path(path)
    if not path.is_file():
        return {"path": str(path), "state": "missing"}
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(block)
    return {"path": str(path), "state": "present", "bytes": path.stat().st_size,
            "sha256": hasher.hexdigest()}


def evaluator_identity(root: Path | None = None) -> dict:
    root = Path(root).resolve() if root is not None else Path(__file__).resolve().parent
    # Hash executed source bytes, not Git state or a mutable version label.
    files = {path.relative_to(root).as_posix(): file_identity(path)["sha256"]
             for path in sorted(root.rglob("*.py"))
             if "tests" not in path.relative_to(root).parts and "__pycache__" not in path.parts}
    manifest = json.dumps(files, sort_keys=True, separators=(",", ":")).encode("utf-8")
    version = root.parent / "version.json"
    return {"root": str(root), "source_sha256": hashlib.sha256(manifest).hexdigest(),
            "files": files, "version_manifest": file_identity(version)}


def runtime_environment() -> dict:
    packages = {}
    for name in ("pywin32", "lxml", "Pillow", "psutil"):
        try:
            packages[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            packages[name] = None
    environment = {"platform": platform.platform(), "machine": platform.machine(),
                   "python": platform.python_version(), "python_executable": sys.executable,
                   "packages": packages, "office": {}, "certified": False}
    if sys.platform == "win32":
        import winreg
        # Installation metadata is a diagnostic identity, not a compatibility
        # certification or a claim about a successfully opened COM instance.
        for executable in ("WINWORD.EXE", "EXCEL.EXE"):
            try:
                key_name = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{executable}"
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_name) as key:
                    path = Path(winreg.QueryValue(key, None).strip('"'))
                identity = file_identity(path)
                try:
                    import win32api
                    info = win32api.GetFileVersionInfo(str(path), "\\")
                    high, low = info["FileVersionMS"], info["FileVersionLS"]
                    identity["file_version"] = ".".join(map(str, (high >> 16, high & 65535, low >> 16, low & 65535)))
                except Exception as exc:
                    identity["version_read_error"] = str(exc)
                environment["office"][executable] = identity
            except OSError as exc:
                environment["office"][executable] = {"state": "unavailable", "error": str(exc)}
    return environment
