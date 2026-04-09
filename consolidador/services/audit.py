from __future__ import annotations

from datetime import datetime, timedelta

from ..db import execute, fetch_all


def log_event(
    event_type: str,
    details: str = "",
    *,
    actor: str = "",
    remote_addr: str = "",
    submission_id: str = "",
    file_name: str = "",
) -> None:
    execute(
        """
        INSERT INTO audit_logs (
            created_at, event_type, actor, remote_addr, details, submission_id, file_name
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            event_type,
            actor,
            remote_addr,
            details,
            submission_id,
            file_name,
        ),
    )


def count_recent_events(event_type: str, remote_addr: str, minutes: int) -> int:
    cutoff = (datetime.now() - timedelta(minutes=minutes)).strftime("%Y-%m-%d %H:%M:%S")
    row = fetch_all(
        """
        SELECT id
        FROM audit_logs
        WHERE event_type = ? AND remote_addr = ? AND created_at >= ?
        """,
        (event_type, remote_addr, cutoff),
    )
    return len(row)
