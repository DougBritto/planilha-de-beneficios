from __future__ import annotations

from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode

from flask import Blueprint, current_app, flash, redirect, render_template, request, send_file, session, url_for

from ..config import AppSettings
from ..plan_types import MAIN_SHEET_NAMES, OUTPUT_BASENAMES, PLAN_LABELS, is_valid_plan_type
from ..security import admin_required, is_admin_password_valid
from ..services.audit import count_recent_events, log_event
from ..services.consolidation import build_consolidated_workbook, write_consolidation_output_for_storage
from ..services.repository import (
    delete_upload,
    delete_uploads_by_submission,
    fetch_audit_logs_filtered,
    fetch_consolidation,
    fetch_consolidations,
    fetch_dashboard_stats,
    fetch_submission_group,
    fetch_submission_groups,
    fetch_upload,
    fetch_uploads,
    insert_consolidation_record,
)
from ..services.secure_storage import build_download_response
from werkzeug.utils import secure_filename

bp = Blueprint("admin", __name__, url_prefix="/admin")
ADMIN_VIEWS = {
    "dashboard": "admin.dashboard",
    "uploads": "admin.uploads_page",
    "audit": "admin.audit_page",
}


def _settings() -> AppSettings:
    return current_app.config["APP_SETTINGS"]


def _remote_addr() -> str:
    forwarded = request.headers.get("X-Forwarded-For", "")
    candidate = forwarded.split(",")[0].strip() if forwarded else (request.remote_addr or "").strip()
    return candidate[:120]


def _clean_plan_type(value: str | None) -> str:
    candidate = (value or "").strip().lower()
    return candidate if is_valid_plan_type(candidate) else ""


def _clean_return_view(value: str | None) -> str:
    candidate = (value or "").strip().lower()
    return candidate if candidate in ADMIN_VIEWS else "dashboard"


def _admin_url(
    *,
    view_name: str = "dashboard",
    plan_type_filter: str = "",
    submission_id: str = "",
    sender_query: str = "",
    date_from: str = "",
    date_to: str = "",
) -> str:
    params = {
        key: value
        for key, value in {
            "plan_type_filter": plan_type_filter,
            "submission_id": submission_id,
            "sender_query": sender_query,
            "date_from": date_from,
            "date_to": date_to,
        }.items()
        if value
    }
    base_url = url_for(ADMIN_VIEWS[_clean_return_view(view_name)])
    if not params:
        return base_url
    return f"{base_url}?{urlencode(params)}"


def _admin_redirect(
    *,
    view_name: str = "dashboard",
    plan_type_filter: str = "",
    submission_id: str = "",
    sender_query: str = "",
    date_from: str = "",
    date_to: str = "",
):
    return redirect(
        _admin_url(
            view_name=view_name,
            plan_type_filter=plan_type_filter,
            submission_id=submission_id,
            sender_query=sender_query,
            date_from=date_from,
            date_to=date_to,
        )
    )


def _read_filters(source) -> dict[str, str]:
    return {
        "plan_type_filter": _clean_plan_type(source.get("plan_type_filter")),
        "submission_id": (source.get("submission_id") or "").strip(),
        "sender_query": (source.get("sender_query") or "").strip(),
        "date_from": _clean_date(source.get("date_from")),
        "date_to": _clean_date(source.get("date_to")),
    }


def _apply_filter_window(filters: dict[str, str]) -> tuple[str | None, str | None]:
    submitted_from = _date_floor(filters["date_from"]) if filters["date_from"] else None
    submitted_to = _date_ceiling(filters["date_to"]) if filters["date_to"] else None
    return submitted_from, submitted_to


def _admin_base_context(*, active_admin_page: str, filters: dict[str, str]) -> dict[str, object]:
    selected_submission = fetch_submission_group(filters["submission_id"]) if filters["submission_id"] else None
    return {
        "active_admin_page": active_admin_page,
        "plan_type_filter": filters["plan_type_filter"],
        "sender_query": filters["sender_query"],
        "date_from": filters["date_from"],
        "date_to": filters["date_to"],
        "selected_submission_id": filters["submission_id"],
        "selected_submission": selected_submission,
    }


def _clean_date(value: str | None) -> str:
    candidate = (value or "").strip()
    if not candidate:
        return ""
    try:
        datetime.strptime(candidate, "%Y-%m-%d")
    except ValueError:
        return ""
    return candidate


def _date_floor(value: str) -> str:
    return f"{value} 00:00:00"


def _date_ceiling(value: str) -> str:
    return f"{value} 23:59:59"


def _cleanup_upload_file(path: Path, upload_root: Path) -> None:
    try:
        resolved_root = upload_root.resolve()
        resolved_path = path.resolve(strict=False)
    except Exception:
        return

    if not resolved_path.is_relative_to(resolved_root):
        return

    path.unlink(missing_ok=True)

    for parent in path.parents:
        if parent == resolved_root.parent:
            break
        if not parent.is_relative_to(resolved_root):
            break
        try:
            parent.rmdir()
        except OSError:
            break


def _safe_download_name(filename: str, fallback: str) -> str:
    candidate = secure_filename((filename or "").strip())
    return candidate or fallback


@bp.get("/login")
def login():
    if session.get("admin_authenticated"):
        return redirect(url_for("admin.dashboard"))
    return render_template("admin_login.html")


@bp.post("/login")
def login_post():
    settings = _settings()
    remote_addr = _remote_addr()
    failed_attempts = count_recent_events(
        "admin_login_failed",
        remote_addr,
        settings.login_lock_minutes,
    )
    if failed_attempts >= settings.login_max_attempts:
        flash(
            (
                "Acesso temporariamente bloqueado por excesso de tentativas. "
                f"Tente novamente em {settings.login_lock_minutes} minuto(s)."
            ),
            "error",
        )
        return redirect(url_for("admin.login"))

    password = request.form.get("password", "")
    if is_admin_password_valid(password):
        session.clear()
        session["admin_authenticated"] = True
        session["admin_authenticated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        session["admin_last_seen"] = datetime.now().isoformat(timespec="seconds")
        log_event("admin_login_success", "Login administrativo realizado.", actor="admin", remote_addr=remote_addr)
        return redirect(url_for("admin.dashboard"))

    log_event("admin_login_failed", "Senha administrativa invalida.", actor="admin", remote_addr=remote_addr)
    flash("Senha invalida.", "error")
    return redirect(url_for("admin.login"))


@bp.get("/logout")
def logout():
    if session.get("admin_authenticated"):
        log_event("admin_logout", "Sessao administrativa encerrada.", actor="admin", remote_addr=_remote_addr())
    session.clear()
    flash("Sessao encerrada.", "success")
    return redirect(url_for("admin.login"))


@bp.get("")
@admin_required
def dashboard():
    filters = _read_filters(request.args)
    submitted_from, submitted_to = _apply_filter_window(filters)

    stats = fetch_dashboard_stats(
        plan_type=filters["plan_type_filter"] or None,
        submission_id=filters["submission_id"] or None,
        sender_query=filters["sender_query"] or None,
        submitted_from=submitted_from,
        submitted_to=submitted_to,
    )
    recent_submissions = fetch_submission_groups(
        limit=20,
        plan_type=filters["plan_type_filter"] or None,
        sender_query=filters["sender_query"] or None,
        submitted_from=submitted_from,
        submitted_to=submitted_to,
    )
    recent_outputs = fetch_consolidations(
        limit=10,
        plan_type=filters["plan_type_filter"] or None,
        submission_id=filters["submission_id"] or None,
        created_from=submitted_from,
        created_to=submitted_to,
    )
    settings = _settings()

    return render_template(
        "admin_dashboard_v2.html",
        stats=stats,
        recent_submissions=recent_submissions,
        recent_outputs=recent_outputs,
        default_sheet_name=settings.default_sheet_name,
        block_duplicate_files=settings.block_duplicate_files,
        health_template_available=bool(settings.health_template_path and settings.health_template_path.exists()),
        dental_template_available=bool(settings.dental_template_path and settings.dental_template_path.exists()),
        **_admin_base_context(active_admin_page="dashboard", filters=filters),
    )


@bp.get("/uploads")
@admin_required
def uploads_page():
    filters = _read_filters(request.args)
    submitted_from, submitted_to = _apply_filter_window(filters)
    recent_files = fetch_uploads(
        order_desc=True,
        limit=150,
        plan_type=filters["plan_type_filter"] or None,
        submission_id=filters["submission_id"] or None,
        sender_query=filters["sender_query"] or None,
        submitted_from=submitted_from,
        submitted_to=submitted_to,
    )

    return render_template(
        "admin_uploads.html",
        recent_files=recent_files,
        **_admin_base_context(active_admin_page="uploads", filters=filters),
    )


@bp.get("/audit")
@admin_required
def audit_page():
    filters = _read_filters(request.args)
    submitted_from, submitted_to = _apply_filter_window(filters)
    recent_audit_logs = fetch_audit_logs_filtered(
        limit=80,
        submission_id=filters["submission_id"] or None,
        query=filters["sender_query"] or None,
        created_from=submitted_from,
        created_to=submitted_to,
    )

    return render_template(
        "admin_audit.html",
        recent_audit_logs=recent_audit_logs,
        **_admin_base_context(active_admin_page="audit", filters=filters),
    )


@bp.post("/consolidate")
@admin_required
def consolidate():
    settings = _settings()
    plan_type = (request.form.get("plan_type") or "").strip().lower()
    filters = _read_filters(request.form)
    if not is_valid_plan_type(plan_type):
        flash("Selecione um tipo valido de consolidacao.", "error")
        return _admin_redirect(view_name="dashboard", **filters)

    sheet_name = (request.form.get("sheet_name") or "").strip() or settings.default_sheet_name or None
    remove_duplicates = request.form.get("remove_duplicates") == "on"
    ignore_empty = request.form.get("ignore_empty") != "off"
    scoped_submission_id = filters["submission_id"]

    files = fetch_uploads(
        order_desc=False,
        plan_type=plan_type,
        submission_id=scoped_submission_id or None,
    )
    if not files:
        suffix = f" no protocolo {scoped_submission_id}" if scoped_submission_id else ""
        flash(f"Ainda nao ha arquivos {PLAN_LABELS[plan_type]} enviados para consolidar{suffix}.", "error")
        return _admin_redirect(view_name="dashboard", **filters)

    result = build_consolidated_workbook(
        files,
        plan_type=plan_type,
        sheet_name=sheet_name,
        base_template_path=settings.base_template_for(plan_type),
        remove_duplicates=remove_duplicates,
        ignore_empty=ignore_empty,
        include_source_columns=settings.include_source_columns,
    )

    if result.consolidated_rows == 0 and result.invalid_files == 0:
        flash(f"Nenhum dado valido foi encontrado para consolidar em {PLAN_LABELS[plan_type]}.", "error")
        return _admin_redirect(view_name="dashboard", **filters)

    output_name = f"{OUTPUT_BASENAMES[plan_type]}.xlsx"
    stored_output_name = f"{OUTPUT_BASENAMES[plan_type]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = settings.output_dir / stored_output_name
    try:
        stored_output_path = write_consolidation_output_for_storage(
            result,
            output_path,
            main_sheet_name=MAIN_SHEET_NAMES[plan_type],
            base_template_path=settings.base_template_for(plan_type),
            sheet_name=sheet_name,
        )
    except Exception as exc:
        current_app.logger.exception("Falha ao gerar o consolidado %s.", PLAN_LABELS[plan_type])
        log_event(
            "consolidation_generation_failed",
            f"Erro ao gerar consolidado {PLAN_LABELS[plan_type]}: {exc}",
            actor="admin",
            remote_addr=_remote_addr(),
        )
        flash(
            (
                f"Nao foi possivel gerar o consolidado {PLAN_LABELS[plan_type]}. "
                "Revise as inconsistencias e tente novamente."
            ),
            "error",
        )
        return _admin_redirect(view_name="dashboard", **filters)

    insert_consolidation_record(
        {
            "plan_type": plan_type,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "output_name": output_name,
            "stored_path": str(stored_output_path),
            "scope_submission_id": scoped_submission_id,
            "sheet_name": sheet_name or "",
            "remove_duplicates": int(remove_duplicates),
            "ignore_empty": int(ignore_empty),
            "total_files": len(files),
            "valid_files": result.valid_files,
            "invalid_files": result.invalid_files,
            "rows_generated": result.consolidated_rows,
            "requested_by": "admin",
        }
    )
    log_event(
        "consolidation_generated",
        (
            f"Consolidado {PLAN_LABELS[plan_type]} gerado com {result.consolidated_rows} linha(s), "
            f"{result.valid_files} arquivo(s) validos e {result.invalid_files} inconsistente(s)."
            + (f" Escopo: protocolo {scoped_submission_id}." if scoped_submission_id else "")
        ),
        actor="admin",
        remote_addr=_remote_addr(),
        file_name=output_name,
        submission_id=scoped_submission_id,
    )

    return build_download_response(
        stored_output_path,
        download_name=_safe_download_name(output_name, f"consolidado_{plan_type}.xlsx"),
        fallback_mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@bp.get("/uploads/<int:upload_id>/download")
@admin_required
def download_upload(upload_id: int):
    upload = fetch_upload(upload_id)
    if upload is None or not Path(upload.stored_path).exists():
        flash("Arquivo nao encontrado.", "error")
        return redirect(url_for("admin.dashboard"))
    log_event(
        "upload_downloaded",
        f"Arquivo {upload.original_name} baixado pelo admin.",
        actor="admin",
        remote_addr=_remote_addr(),
        submission_id=upload.submission_id,
        file_name=upload.original_name,
    )
    return build_download_response(
        Path(upload.stored_path),
        download_name=_safe_download_name(upload.original_name, f"upload_{upload.id}"),
        fallback_mimetype=upload.content_type or "application/octet-stream",
    )


@bp.post("/uploads/<int:upload_id>/delete")
@admin_required
def delete_upload_route(upload_id: int):
    upload = fetch_upload(upload_id)
    filters = _read_filters(request.form)
    return_view = _clean_return_view(request.form.get("return_view"))

    if upload is None:
        flash("Arquivo nao encontrado para exclusao.", "error")
        return _admin_redirect(view_name=return_view, **filters)

    _cleanup_upload_file(Path(upload.stored_path), _settings().upload_dir)
    delete_upload(upload_id)
    log_event(
        "upload_deleted",
        f"Upload {upload.original_name} removido manualmente do painel.",
        actor="admin",
        remote_addr=_remote_addr(),
        submission_id=upload.submission_id,
        file_name=upload.original_name,
    )
    flash(f"Arquivo {upload.original_name} removido com sucesso.", "success")
    filters["submission_id"] = filters["submission_id"] or upload.submission_id
    return _admin_redirect(view_name=return_view, **filters)


@bp.post("/uploads/delete-selected")
@admin_required
def delete_selected_uploads_route():
    upload_ids = [value for value in request.form.getlist("upload_ids") if value.strip().isdigit()]
    filters = _read_filters(request.form)
    return_view = _clean_return_view(request.form.get("return_view"))

    if not upload_ids:
        flash("Selecione pelo menos um arquivo para excluir.", "error")
        return _admin_redirect(view_name=return_view, **filters)

    deleted_count = 0
    deleted_names: list[str] = []
    for upload_id in upload_ids:
        upload = fetch_upload(int(upload_id))
        if upload is None:
            continue
        _cleanup_upload_file(Path(upload.stored_path), _settings().upload_dir)
        delete_upload(upload.id)
        deleted_count += 1
        deleted_names.append(upload.original_name)
        log_event(
            "upload_deleted",
            f"Upload {upload.original_name} removido manualmente do painel.",
            actor="admin",
            remote_addr=_remote_addr(),
            submission_id=upload.submission_id,
            file_name=upload.original_name,
        )

    if not deleted_count:
        flash("Nenhum dos arquivos selecionados foi encontrado para exclusao.", "error")
    else:
        flash(f"{deleted_count} arquivo(s) removido(s) com sucesso.", "success")

    return _admin_redirect(
        view_name=return_view,
        **filters,
    )


@bp.post("/submissions/<submission_id>/delete")
@admin_required
def delete_submission_route(submission_id: str):
    filters = _read_filters(request.form)
    return_view = _clean_return_view(request.form.get("return_view"))
    uploads = fetch_uploads(order_desc=False, submission_id=submission_id)

    if not uploads:
        flash("Lote nao encontrado para exclusao.", "error")
        filters["submission_id"] = ""
        return _admin_redirect(view_name=return_view, **filters)

    for upload in uploads:
        _cleanup_upload_file(Path(upload.stored_path), _settings().upload_dir)

    delete_uploads_by_submission(submission_id)
    log_event(
        "submission_deleted",
        f"Lote {submission_id} removido manualmente com {len(uploads)} arquivo(s).",
        actor="admin",
        remote_addr=_remote_addr(),
        submission_id=submission_id,
    )
    flash(f"Lote {submission_id} removido com {len(uploads)} arquivo(s).", "success")
    filters["submission_id"] = ""
    return _admin_redirect(view_name=return_view, **filters)


@bp.get("/outputs/<int:consolidation_id>/download")
@admin_required
def download_output(consolidation_id: int):
    output = fetch_consolidation(consolidation_id)
    if output is None or not Path(output.stored_path).exists():
        flash("Consolidado nao encontrado.", "error")
        return redirect(url_for("admin.dashboard"))
    log_event(
        "consolidation_downloaded",
        f"Consolidado {output.output_name} baixado pelo admin.",
        actor="admin",
        remote_addr=_remote_addr(),
        submission_id=output.scope_submission_id,
        file_name=output.output_name,
    )
    return build_download_response(
        Path(output.stored_path),
        download_name=_safe_download_name(output.output_name, f"consolidado_{consolidation_id}.xlsx"),
        fallback_mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
