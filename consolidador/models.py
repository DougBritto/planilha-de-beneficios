from __future__ import annotations

from dataclasses import dataclass

from .plan_types import plan_label


@dataclass(slots=True)
class SubmissionFile:
    id: int
    plan_type: str
    submission_id: str
    submitted_at: str
    sender_name: str
    sender_email: str
    note: str
    original_name: str
    stored_name: str
    stored_path: str
    size_bytes: int
    file_hash: str
    content_type: str
    remote_addr: str
    user_agent: str

    @property
    def plan_label(self) -> str:
        return plan_label(self.plan_type)


@dataclass(slots=True)
class SubmissionGroup:
    submission_id: str
    submitted_at: str
    sender_name: str
    sender_email: str
    note: str
    file_count: int
    total_size_bytes: int
    saude_count: int
    odonto_count: int

    @property
    def categories_summary(self) -> str:
        parts: list[str] = []
        if self.saude_count:
            parts.append(f"{self.saude_count} Saude")
        if self.odonto_count:
            parts.append(f"{self.odonto_count} Odonto")
        return " / ".join(parts) if parts else "-"


@dataclass(slots=True)
class ConsolidationRecord:
    id: int
    plan_type: str
    created_at: str
    output_name: str
    stored_path: str
    scope_submission_id: str
    sheet_name: str
    remove_duplicates: int
    ignore_empty: int
    total_files: int
    valid_files: int
    invalid_files: int
    rows_generated: int
    requested_by: str

    @property
    def plan_label(self) -> str:
        return plan_label(self.plan_type)

    @property
    def scope_label(self) -> str:
        return self.scope_submission_id or "Todos os lotes"


@dataclass(slots=True)
class AuditLogEntry:
    id: int
    created_at: str
    event_type: str
    actor: str
    remote_addr: str
    details: str
    submission_id: str
    file_name: str
