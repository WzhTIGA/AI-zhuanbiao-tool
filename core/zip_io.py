from __future__ import annotations

import io
import zipfile
import rarfile
from pathlib import PurePosixPath


def read_xlsx_files_from_archive(archive_bytes: bytes) -> dict[str, bytes]:
    # Check magic bytes for RAR
    # RAR4: b'Rar!\x1a\x07\x00', RAR5: b'Rar!\x1a\x07\x01\x00'
    if archive_bytes.startswith(b'Rar!'):
        return _read_xlsx_files_from_rar(archive_bytes)
    return _read_xlsx_files_from_zip(archive_bytes)


def _read_xlsx_files_from_rar(rar_bytes: bytes) -> dict[str, bytes]:
    out: dict[str, bytes] = {}
    with rarfile.RarFile(io.BytesIO(rar_bytes)) as rf:
        for info in rf.infolist():
            if info.isdir():
                continue
            name = info.filename
            p = PurePosixPath(name.replace('\\', '/'))
            if p.is_absolute() or ".." in p.parts:
                continue
            if p.name.startswith("~$"):
                continue
            if not p.name.lower().endswith(".xlsx"):
                continue
            out[p.name] = rf.read(info)
    return out


def _read_xlsx_files_from_zip(zip_bytes: bytes) -> dict[str, bytes]:
    out: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename
            p = PurePosixPath(name.replace('\\', '/'))
            if p.is_absolute() or ".." in p.parts:
                continue
            if p.name.startswith("~$"):
                continue
            if not p.name.lower().endswith(".xlsx"):
                continue
            out[p.name] = zf.read(info)
    return out


def write_zip(files: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    return buf.getvalue()

