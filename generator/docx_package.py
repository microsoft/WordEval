"""Write OOXML packages atomically without predictable temporary paths."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Mapping, Union
from zipfile import ZIP_DEFLATED, ZipFile


def write_docx_package(files: Mapping[str, Union[bytes, str]], target: Union[str, Path]) -> Path:
    """Write ``files`` as a ZIP package and atomically replace ``target``.

    The package is first written to a uniquely created temporary file in the
    target directory, then moved over ``target`` with ``os.replace``.  Writing
    in the same directory keeps the final step atomic on the same volume and
    avoids a predictable system temp path.
    """
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(
        dir=target.parent, prefix=f".{target.stem}.", suffix=".docx.tmp", delete=False
    )
    temp_path = Path(handle.name)
    try:
        with handle, ZipFile(handle, "w", ZIP_DEFLATED) as zout:
            for name, content in files.items():
                zout.writestr(name, content)
        os.replace(temp_path, target)
    except BaseException:
        temp_path.unlink(missing_ok=True)
        raise
    return target
