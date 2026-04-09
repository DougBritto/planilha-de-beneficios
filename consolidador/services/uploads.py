from __future__ import annotations

import hashlib
import re
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from flask import current_app
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from ..config import AppSettings
from ..plan_types import PLAN_TYPES, PLAN_LABELS, is_valid_plan_type
from .audit import log_event
from .repository import file_hash_exists, insert_upload_record
from .secure_storage import persist_file_for_storage


EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class ValidationError(Exception):
    pass


@dataclass(slots=True)
class UploadBatchResult:
    submission_id: str
    submitted_at: str
    saved_counts: dict[str, int]
    duplicate_files: dict[str, list[str]]

    @property
    def saved_count(self) -> int:
        return sum(self.saved_counts.values())


def _settings() -> AppSettings:
    return current_app.config["APP_SETTINGS"]


def _clean_text(value: str | None, max_length: int) -> str:
    return (value or "").strip()[:max_length]


def validate_sender_data(sender_name: str, sender_email: str, note: str) -> tuple[str, str, str]:
    settings = _settings()
    clean_name = _clean_text(sender_name, settings.max_name_length)
    clean_email = _clean_text(sender_email, settings.max_email_length)
    clean_note = _clean_text(note, settings.max_note_length)

    if clean_email and not EMAIL_PATTERN.match(clean_email):
        raise ValidationError("Informe um e-mail valido ou deixe o campo em branco.")
    return clean_name, clean_email, clean_note


def _allowed_extension(filename: str) -> bool:
    settings = _settings()
    if "." not in filename:
        return False
    return filename.rsplit(".", 1)[1].lower() in settings.allowed_extensions


def validate_uploaded_files(file_groups: dict[str, list[FileStorage]]) -> dict[str, list[FileStorage]]:
    settings = _settings()
    valid_files_by_group: dict[str, list[FileStorage]] = {}
    total_files = 0
    invalid: list[str] = []

    for plan_type, uploaded_files in file_groups.items():
        if not is_valid_plan_type(plan_type):
            continue
        valid_files = [file for file in uploaded_files if file and file.filename]
        valid_files_by_group[plan_type] = valid_files
        total_files += len(valid_files)
        invalid.extend(file.filename for file in valid_files if not _allowed_extension(file.filename))

    if total_files == 0:
        raise ValidationError("Selecione pelo menos um arquivo para continuar.")
    if total_files > settings.max_files_per_upload:
        raise ValidationError(
            f"Envie no maximo {settings.max_files_per_upload} arquivo(s) por vez."
        )

    if invalid:
        invalid_label = ", ".join(invalid)
        raise ValidationError(
            f"Arquivos com extensao nao permitida: {invalid_label}. Envie apenas CSV ou Excel."
        )
    return valid_files_by_group


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def save_uploaded_files(
    file_groups: dict[str, list[FileStorage]],
    *,
    sender_name: str,
    sender_email: str,
    note: str,
    remote_addr: str,
    user_agent: str,
) -> UploadBatchResult:
    settings = _settings()
    valid_files_by_group = validate_uploaded_files(file_groups)
    clean_name, clean_email, clean_note = validate_sender_data(sender_name, sender_email, note)

    submission_id = uuid.uuid4().hex[:10]
    submitted_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    submission_folder = settings.upload_dir / submission_id
    submission_folder.mkdir(parents=True, exist_ok=True)

    saved_counts = {plan_type: 0 for plan_type in PLAN_TYPES}
    duplicate_files = {plan_type: [] for plan_type in PLAN_TYPES}

    for plan_type in PLAN_TYPES:
        typed_files = valid_files_by_group.get(plan_type, [])
        if not typed_files:
            continue

        category_folder = submission_folder / plan_type
        category_folder.mkdir(parents=True, exist_ok=True)

        for index, file in enumerate(typed_files, start=1):
            original_name = file.filename or f"arquivo_{index}"
            safe_name = secure_filename(original_name) or f"arquivo_{index}"
            stored_name = f"{uuid.uuid4().hex[:12]}_{safe_name}"
            plaintext_destination = category_folder / stored_name
            file.save(plaintext_destination)

            file_hash = _hash_file(plaintext_destination)
            original_size_bytes = plaintext_destination.stat().st_size
            if settings.block_duplicate_files and file_hash_exists(file_hash):
                plaintext_destination.unlink(missing_ok=True)
                duplicate_files[plan_type].append(original_name)
                log_event(
                    "upload_duplicate_blocked",
                    f"Arquivo {PLAN_LABELS[plan_type]} bloqueado por hash duplicado.",
                    actor=clean_name or clean_email or "publico",
                    remote_addr=remote_addr,
                    submission_id=submission_id,
                    file_name=original_name,
                )
                continue

            destination = persist_file_for_storage(plaintext_destination, plaintext_destination)

            insert_upload_record(
                {
                    "plan_type": plan_type,
                    "submission_id": submission_id,
                    "submitted_at": submitted_at,
                    "sender_name": clean_name,
                    "sender_email": clean_email,
                    "note": clean_note,
                    "original_name": original_name,
                    "stored_name": destination.name,
                    "stored_path": str(destination),
                    "size_bytes": original_size_bytes,
                    "file_hash": file_hash,
                    "content_type": file.mimetype or "",
                    "remote_addr": remote_addr,
                    "user_agent": user_agent[:255],
                }
            )
            log_event(
                "upload_saved",
                f"Arquivo {PLAN_LABELS[plan_type]} recebido com sucesso.",
                actor=clean_name or clean_email or "publico",
                remote_addr=remote_addr,
                submission_id=submission_id,
                file_name=original_name,
            )
            saved_counts[plan_type] += 1

    if sum(saved_counts.values()) == 0 and any(duplicate_files.values()):
        log_event(
            "upload_batch_without_new_files",
            "Nenhum arquivo novo foi salvo no lote.",
            actor=clean_name or clean_email or "publico",
            remote_addr=remote_addr,
            submission_id=submission_id,
        )

    return UploadBatchResult(
        submission_id=submission_id,
        submitted_at=submitted_at,
        saved_counts=saved_counts,
        duplicate_files=duplicate_files,
    )
