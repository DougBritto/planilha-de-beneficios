from __future__ import annotations

from ..db import execute, fetch_all, fetch_one
from ..models import AuditLogEntry, ConsolidationRecord, SubmissionFile, SubmissionGroup


def row_to_submission_file(row) -> SubmissionFile:
    payload = dict(row)
    for key in ["plan_type", "sender_name", "sender_email", "note", "stored_name", "file_hash", "content_type", "remote_addr", "user_agent"]:
        payload[key] = payload.get(key) or ""
    return SubmissionFile(**payload)


def row_to_submission_group(row) -> SubmissionGroup:
    payload = dict(row)
    for key in ["sender_name", "sender_email", "note"]:
        payload[key] = payload.get(key) or ""
    for key in ["saude_count", "odonto_count"]:
        payload[key] = int(payload.get(key) or 0)
    return SubmissionGroup(**payload)


def row_to_consolidation_record(row) -> ConsolidationRecord:
    payload = dict(row)
    for key in ["plan_type", "scope_submission_id", "sheet_name", "requested_by"]:
        payload[key] = payload.get(key) or ""
    return ConsolidationRecord(**payload)


def row_to_audit_log(row) -> AuditLogEntry:
    payload = dict(row)
    for key in ["actor", "remote_addr", "details", "submission_id", "file_name"]:
        payload[key] = payload.get(key) or ""
    return AuditLogEntry(**payload)


def insert_upload_record(payload: dict[str, object]) -> int:
    return execute(
        """
        INSERT INTO uploads (
            plan_type, submission_id, submitted_at, sender_name, sender_email, note,
            original_name, stored_name, stored_path, size_bytes, file_hash,
            content_type, remote_addr, user_agent
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            payload["plan_type"],
            payload["submission_id"],
            payload["submitted_at"],
            payload["sender_name"],
            payload["sender_email"],
            payload["note"],
            payload["original_name"],
            payload["stored_name"],
            payload["stored_path"],
            payload["size_bytes"],
            payload["file_hash"],
            payload["content_type"],
            payload["remote_addr"],
            payload["user_agent"],
        ),
    )


def file_hash_exists(file_hash: str) -> bool:
    row = fetch_one("SELECT id FROM uploads WHERE file_hash = ? LIMIT 1", (file_hash,))
    return row is not None


def fetch_upload(upload_id: int) -> SubmissionFile | None:
    row = fetch_one(
        """
        SELECT id, COALESCE(plan_type, '') AS plan_type, submission_id, submitted_at, sender_name, sender_email, note,
               original_name, COALESCE(stored_name, '') AS stored_name, stored_path,
               size_bytes, COALESCE(file_hash, '') AS file_hash,
               COALESCE(content_type, '') AS content_type,
               COALESCE(remote_addr, '') AS remote_addr,
               COALESCE(user_agent, '') AS user_agent
        FROM uploads
        WHERE id = ?
        """,
        (upload_id,),
    )
    return row_to_submission_file(row) if row else None


def fetch_uploads(
    order_desc: bool = True,
    limit: int | None = None,
    plan_type: str | None = None,
    submission_id: str | None = None,
    sender_query: str | None = None,
    submitted_from: str | None = None,
    submitted_to: str | None = None,
) -> list[SubmissionFile]:
    query = """
        SELECT id, COALESCE(plan_type, '') AS plan_type, submission_id, submitted_at, sender_name, sender_email, note,
               original_name, COALESCE(stored_name, '') AS stored_name, stored_path,
               size_bytes, COALESCE(file_hash, '') AS file_hash,
               COALESCE(content_type, '') AS content_type,
               COALESCE(remote_addr, '') AS remote_addr,
               COALESCE(user_agent, '') AS user_agent
        FROM uploads
    """
    params: list[object] = []
    conditions: list[str] = []
    if plan_type:
        conditions.append("plan_type = ?")
        params.append(plan_type)
    if submission_id:
        conditions.append("submission_id = ?")
        params.append(submission_id)
    if sender_query:
        conditions.append("(sender_name LIKE ? OR sender_email LIKE ? OR original_name LIKE ?)")
        token = f"%{sender_query.strip()}%"
        params.extend((token, token, token))
    if submitted_from:
        conditions.append("submitted_at >= ?")
        params.append(submitted_from)
    if submitted_to:
        conditions.append("submitted_at <= ?")
        params.append(submitted_to)
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += " ORDER BY submitted_at {direction}, id {direction}".format(direction="DESC" if order_desc else "ASC")
    if limit is not None:
        query += f" LIMIT {int(limit)}"
    return [row_to_submission_file(row) for row in fetch_all(query, tuple(params))]


def fetch_submission_group(submission_id: str) -> SubmissionGroup | None:
    row = fetch_one(
        """
        SELECT
            submission_id,
            MAX(submitted_at) AS submitted_at,
            COALESCE(MAX(NULLIF(sender_name, '')), '') AS sender_name,
            COALESCE(MAX(NULLIF(sender_email, '')), '') AS sender_email,
            COALESCE(MAX(NULLIF(note, '')), '') AS note,
            COUNT(*) AS file_count,
            COALESCE(SUM(size_bytes), 0) AS total_size_bytes,
            COALESCE(SUM(CASE WHEN plan_type = 'saude' THEN 1 ELSE 0 END), 0) AS saude_count,
            COALESCE(SUM(CASE WHEN plan_type = 'odonto' THEN 1 ELSE 0 END), 0) AS odonto_count
        FROM uploads
        WHERE submission_id = ?
        GROUP BY submission_id
        """,
        (submission_id,),
    )
    return row_to_submission_group(row) if row else None


def fetch_submission_groups(
    limit: int = 100,
    plan_type: str | None = None,
    sender_query: str | None = None,
    submitted_from: str | None = None,
    submitted_to: str | None = None,
) -> list[SubmissionGroup]:
    where_parts: list[str] = []
    params: list[object] = []
    if plan_type:
        where_parts.append("plan_type = ?")
        params.append(plan_type)
    if sender_query:
        where_parts.append("(sender_name LIKE ? OR sender_email LIKE ? OR note LIKE ? OR submission_id LIKE ?)")
        token = f"%{sender_query.strip()}%"
        params.extend((token, token, token, token))
    if submitted_from:
        where_parts.append("submitted_at >= ?")
        params.append(submitted_from)
    if submitted_to:
        where_parts.append("submitted_at <= ?")
        params.append(submitted_to)

    where_clause = ""
    if where_parts:
        where_clause = "WHERE " + " AND ".join(where_parts)

    rows = fetch_all(
        """
        SELECT
            submission_id,
            MAX(submitted_at) AS submitted_at,
            COALESCE(MAX(NULLIF(sender_name, '')), '') AS sender_name,
            COALESCE(MAX(NULLIF(sender_email, '')), '') AS sender_email,
            COALESCE(MAX(NULLIF(note, '')), '') AS note,
            COUNT(*) AS file_count,
            COALESCE(SUM(size_bytes), 0) AS total_size_bytes,
            COALESCE(SUM(CASE WHEN plan_type = 'saude' THEN 1 ELSE 0 END), 0) AS saude_count,
            COALESCE(SUM(CASE WHEN plan_type = 'odonto' THEN 1 ELSE 0 END), 0) AS odonto_count
        FROM uploads
        """
        + where_clause
        + f"""
        GROUP BY submission_id
        ORDER BY MAX(submitted_at) DESC
        LIMIT {int(limit)}
        """,
        tuple(params),
    )
    return [row_to_submission_group(row) for row in rows]


def fetch_dashboard_stats(
    plan_type: str | None = None,
    submission_id: str | None = None,
    sender_query: str | None = None,
    submitted_from: str | None = None,
    submitted_to: str | None = None,
) -> dict[str, int]:
    upload_conditions: list[str] = []
    upload_params: list[object] = []
    if plan_type:
        upload_conditions.append("plan_type = ?")
        upload_params.append(plan_type)
    if submission_id:
        upload_conditions.append("submission_id = ?")
        upload_params.append(submission_id)
    if sender_query:
        token = f"%{sender_query.strip()}%"
        upload_conditions.append("(sender_name LIKE ? OR sender_email LIKE ? OR note LIKE ? OR original_name LIKE ?)")
        upload_params.extend((token, token, token, token))
    if submitted_from:
        upload_conditions.append("submitted_at >= ?")
        upload_params.append(submitted_from)
    if submitted_to:
        upload_conditions.append("submitted_at <= ?")
        upload_params.append(submitted_to)

    upload_where = ""
    if upload_conditions:
        upload_where = " WHERE " + " AND ".join(upload_conditions)

    output_conditions: list[str] = []
    output_params: list[object] = []
    if plan_type:
        output_conditions.append("plan_type = ?")
        output_params.append(plan_type)
    if submission_id:
        output_conditions.append("scope_submission_id = ?")
        output_params.append(submission_id)
    if submitted_from:
        output_conditions.append("created_at >= ?")
        output_params.append(submitted_from)
    if submitted_to:
        output_conditions.append("created_at <= ?")
        output_params.append(submitted_to)

    output_where = ""
    if output_conditions:
        output_where = " WHERE " + " AND ".join(output_conditions)

    total_files = fetch_one("SELECT COUNT(*) AS total FROM uploads" + upload_where, tuple(upload_params))
    total_submissions = fetch_one(
        "SELECT COUNT(DISTINCT submission_id) AS total FROM uploads" + upload_where,
        tuple(upload_params),
    )
    total_outputs = fetch_one("SELECT COUNT(*) AS total FROM consolidations" + output_where, tuple(output_params))
    total_health_files = fetch_one(
        "SELECT COUNT(*) AS total FROM uploads"
        + upload_where
        + (" AND plan_type = 'saude'" if upload_where else " WHERE plan_type = 'saude'"),
        tuple(upload_params),
    )
    total_dental_files = fetch_one(
        "SELECT COUNT(*) AS total FROM uploads"
        + upload_where
        + (" AND plan_type = 'odonto'" if upload_where else " WHERE plan_type = 'odonto'"),
        tuple(upload_params),
    )
    return {
        "total_files": int(total_files["total"]) if total_files else 0,
        "total_submissions": int(total_submissions["total"]) if total_submissions else 0,
        "total_outputs": int(total_outputs["total"]) if total_outputs else 0,
        "total_saude_files": int(total_health_files["total"]) if total_health_files else 0,
        "total_odonto_files": int(total_dental_files["total"]) if total_dental_files else 0,
    }


def insert_consolidation_record(payload: dict[str, object]) -> int:
    return execute(
        """
        INSERT INTO consolidations (
            plan_type, created_at, output_name, stored_path, scope_submission_id, sheet_name, remove_duplicates, ignore_empty,
            total_files, valid_files, invalid_files, rows_generated, requested_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            payload["plan_type"],
            payload["created_at"],
            payload["output_name"],
            payload["stored_path"],
            payload["scope_submission_id"],
            payload["sheet_name"],
            payload["remove_duplicates"],
            payload["ignore_empty"],
            payload["total_files"],
            payload["valid_files"],
            payload["invalid_files"],
            payload["rows_generated"],
            payload["requested_by"],
        ),
    )


def fetch_consolidation(consolidation_id: int) -> ConsolidationRecord | None:
    row = fetch_one(
        """
        SELECT id, COALESCE(plan_type, '') AS plan_type, created_at, output_name, stored_path,
               COALESCE(scope_submission_id, '') AS scope_submission_id, COALESCE(sheet_name, '') AS sheet_name,
               remove_duplicates, ignore_empty, total_files, valid_files, invalid_files,
               rows_generated, COALESCE(requested_by, '') AS requested_by
        FROM consolidations
        WHERE id = ?
        """,
        (consolidation_id,),
    )
    return row_to_consolidation_record(row) if row else None


def fetch_consolidations(
    limit: int = 20,
    plan_type: str | None = None,
    submission_id: str | None = None,
    created_from: str | None = None,
    created_to: str | None = None,
) -> list[ConsolidationRecord]:
    query = """
        SELECT id, COALESCE(plan_type, '') AS plan_type, created_at, output_name, stored_path,
               COALESCE(scope_submission_id, '') AS scope_submission_id, COALESCE(sheet_name, '') AS sheet_name,
               remove_duplicates, ignore_empty, total_files, valid_files, invalid_files,
               rows_generated, COALESCE(requested_by, '') AS requested_by
        FROM consolidations
    """
    params: list[object] = []
    conditions: list[str] = []
    if plan_type:
        conditions.append("plan_type = ?")
        params.append(plan_type)
    if submission_id:
        conditions.append("scope_submission_id = ?")
        params.append(submission_id)
    if created_from:
        conditions.append("created_at >= ?")
        params.append(created_from)
    if created_to:
        conditions.append("created_at <= ?")
        params.append(created_to)
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += f" ORDER BY created_at DESC, id DESC LIMIT {int(limit)}"
    rows = fetch_all(query, tuple(params))
    return [row_to_consolidation_record(row) for row in rows]


def delete_upload(upload_id: int) -> None:
    execute("DELETE FROM uploads WHERE id = ?", (upload_id,))


def delete_uploads_by_submission(submission_id: str) -> None:
    execute("DELETE FROM uploads WHERE submission_id = ?", (submission_id,))


def fetch_audit_logs(limit: int = 30) -> list[AuditLogEntry]:
    return fetch_audit_logs_filtered(limit=limit)


def fetch_audit_logs_filtered(
    limit: int = 30,
    submission_id: str | None = None,
    query: str | None = None,
    created_from: str | None = None,
    created_to: str | None = None,
) -> list[AuditLogEntry]:
    sql = """
        SELECT id, created_at, event_type, COALESCE(actor, '') AS actor,
               COALESCE(remote_addr, '') AS remote_addr, COALESCE(details, '') AS details,
               COALESCE(submission_id, '') AS submission_id, COALESCE(file_name, '') AS file_name
        FROM audit_logs
    """
    params: list[object] = []
    conditions: list[str] = []
    if submission_id:
        conditions.append("submission_id = ?")
        params.append(submission_id)
    if query:
        token = f"%{query.strip()}%"
        conditions.append("(event_type LIKE ? OR details LIKE ? OR file_name LIKE ? OR actor LIKE ?)")
        params.extend((token, token, token, token))
    if created_from:
        conditions.append("created_at >= ?")
        params.append(created_from)
    if created_to:
        conditions.append("created_at <= ?")
        params.append(created_to)
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += f" ORDER BY created_at DESC, id DESC LIMIT {int(limit)}"
    rows = fetch_all(sql, tuple(params))
    return [row_to_audit_log(row) for row in rows]
