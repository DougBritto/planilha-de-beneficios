from __future__ import annotations

import shutil
import tempfile
from io import BytesIO
from pathlib import Path

from cryptography.fernet import Fernet
from flask import current_app, send_file

from ..config import AppSettings


def _settings() -> AppSettings:
    return current_app.config["APP_SETTINGS"]


def is_storage_encryption_enabled() -> bool:
    return bool(_settings().data_encryption_key)


def _fernet() -> Fernet:
    key = _settings().data_encryption_key
    if not key:
        raise RuntimeError("A criptografia de armazenamento nao esta configurada.")
    return Fernet(key.encode("utf-8"))


def _encrypted_path(path: Path) -> Path:
    return path.with_name(f"{path.name}.enc")


def persist_file_for_storage(source_path: Path, target_path: Path) -> Path:
    target_path.parent.mkdir(parents=True, exist_ok=True)
    if not is_storage_encryption_enabled():
        if source_path != target_path:
            shutil.move(str(source_path), str(target_path))
        return target_path

    encrypted_path = _encrypted_path(target_path)
    payload = source_path.read_bytes()
    encrypted_payload = _fernet().encrypt(payload)
    encrypted_path.write_bytes(encrypted_payload)
    source_path.unlink(missing_ok=True)
    return encrypted_path


def create_processing_copy(path: Path, *, suffix: str = "") -> Path:
    suffix = suffix or path.suffix or ".tmp"
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    temp_path = Path(temp_file.name)
    temp_file.close()

    if path.suffix.lower() == ".enc":
        decrypted_payload = _fernet().decrypt(path.read_bytes())
        temp_path.write_bytes(decrypted_payload)
        return temp_path

    shutil.copy2(path, temp_path)
    return temp_path


def build_download_response(path: Path, *, download_name: str, fallback_mimetype: str) -> object:
    if path.suffix.lower() == ".enc":
        decrypted_payload = _fernet().decrypt(path.read_bytes())
        return send_file(
            BytesIO(decrypted_payload),
            as_attachment=True,
            download_name=download_name,
            mimetype=fallback_mimetype,
            max_age=0,
        )
    return send_file(path, as_attachment=True, download_name=download_name, mimetype=fallback_mimetype, max_age=0)
