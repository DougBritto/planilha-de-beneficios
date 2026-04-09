from __future__ import annotations

from flask import Blueprint, flash, redirect, render_template, request, url_for

from ..config import AppSettings
from ..plan_types import PLAN_LABELS
from ..services.audit import count_recent_events, log_event
from ..services.uploads import ValidationError, save_uploaded_files

bp = Blueprint("public", __name__)


def _settings() -> AppSettings:
    from flask import current_app

    return current_app.config["APP_SETTINGS"]


def _remote_addr() -> str:
    forwarded = request.headers.get("X-Forwarded-For", "")
    candidate = forwarded.split(",")[0].strip() if forwarded else (request.remote_addr or "").strip()
    return candidate[:120]


@bp.get("/")
def index():
    return render_template("upload.html")


@bp.post("/upload")
def upload():
    settings = _settings()
    remote_addr = _remote_addr()
    recent_uploads = count_recent_events(
        "upload_request_received",
        remote_addr,
        settings.upload_rate_limit_window_minutes,
    )
    if recent_uploads >= settings.upload_rate_limit_count:
        log_event(
            "upload_rate_limited",
            "Tentativa de upload bloqueada por limite de frequencia.",
            actor="publico",
            remote_addr=remote_addr,
        )
        flash(
            (
                "Muitas tentativas de envio em pouco tempo. "
                f"Tente novamente em alguns minutos ({settings.upload_rate_limit_window_minutes} min)."
            ),
            "error",
        )
        return redirect(url_for("public.index"))

    log_event(
        "upload_request_received",
        "Tentativa de upload recebida na pagina publica.",
        actor="publico",
        remote_addr=remote_addr,
    )

    try:
        result = save_uploaded_files(
            {
                "saude": request.files.getlist("files_saude"),
                "odonto": request.files.getlist("files_odonto"),
            },
            sender_name=request.form.get("sender_name", ""),
            sender_email=request.form.get("sender_email", ""),
            note=request.form.get("note", ""),
            remote_addr=remote_addr,
            user_agent=(request.user_agent.string or "")[:255],
        )
    except ValidationError as exc:
        flash(str(exc), "error")
        return redirect(url_for("public.index"))

    if result.saved_count:
        details: list[str] = []
        if result.saved_counts.get("saude"):
            details.append(f"{result.saved_counts['saude']} Saude")
        if result.saved_counts.get("odonto"):
            details.append(f"{result.saved_counts['odonto']} Odonto")
        flash(
            (
                f"Upload concluido. Protocolo {result.submission_id}. "
                f"Recebemos {result.saved_count} arquivo(s): {', '.join(details)}."
            ),
            "success",
        )

    for plan_type, files in result.duplicate_files.items():
        if files:
            flash(
                f"Arquivos {PLAN_LABELS[plan_type]} ignorados por duplicidade: " + ", ".join(files),
                "warning",
            )

    if result.saved_count == 0:
        flash("Nenhum arquivo novo foi salvo neste envio.", "error")

    return redirect(url_for("public.index"))
