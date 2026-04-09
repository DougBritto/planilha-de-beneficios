from __future__ import annotations

from datetime import timedelta

from cryptography.fernet import Fernet
from flask import Flask, has_request_context, render_template, session
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.middleware.proxy_fix import ProxyFix

from .blueprints.admin import bp as admin_bp
from .blueprints.public import bp as public_bp
from .config import AppSettings
from .db import init_app as init_db
from .plan_types import PLAN_LABELS
from .security import get_csrf_token, init_security


def create_app() -> Flask:
    settings = AppSettings.from_env()
    if settings.data_encryption_key:
        Fernet(settings.data_encryption_key.encode("utf-8"))

    app = Flask(__name__, template_folder="templates", static_folder="static")
    app.config.update(
        SECRET_KEY=settings.secret_key,
        MAX_CONTENT_LENGTH=settings.max_content_length_bytes,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_NAME="consolidador_session",
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=settings.session_cookie_secure,
        PERMANENT_SESSION_LIFETIME=timedelta(minutes=settings.session_lifetime_minutes),
    )
    app.config["APP_SETTINGS"] = settings
    app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

    init_db(app)
    init_security(app)

    @app.template_filter("filesize")
    def filesize_filter(num_bytes: int | None) -> str:
        if num_bytes is None:
            return "-"
        value = float(num_bytes)
        for unit in ["B", "KB", "MB", "GB", "TB"]:
            if value < 1024 or unit == "TB":
                return f"{value:.1f} {unit}"
            value /= 1024
        return f"{num_bytes} B"

    @app.template_filter("mask_email")
    def mask_email_filter(value: str | None) -> str:
        email = (value or "").strip()
        if not email or "@" not in email:
            return email or "-"
        local, _, domain = email.partition("@")
        local_mask = local[:2] + "*" * max(0, len(local) - 2) if len(local) > 2 else "*" * len(local)
        domain_name, dot, suffix = domain.partition(".")
        domain_mask = domain_name[:1] + "*" * max(0, len(domain_name) - 1) if domain_name else ""
        suffix_mask = f"{dot}{suffix}" if suffix else ""
        return f"{local_mask}@{domain_mask}{suffix_mask}"

    @app.template_filter("mask_ip")
    def mask_ip_filter(value: str | None) -> str:
        raw = (value or "").strip()
        if not raw:
            return "-"
        if "." in raw:
            parts = raw.split(".")
            if len(parts) == 4:
                return ".".join(parts[:2] + ["***", "***"])
        if ":" in raw:
            parts = raw.split(":")
            if len(parts) > 2:
                return ":".join(parts[:2] + ["****"])
        return raw[:6] + "..." if len(raw) > 6 else raw

    @app.context_processor
    def inject_globals() -> dict[str, object]:
        return {
            "APP_TITLE": settings.app_title,
            "APP_SUBTITLE": settings.app_subtitle,
            "MAX_UPLOAD_MB": settings.max_content_length_mb,
            "MAX_FILES_PER_UPLOAD": settings.max_files_per_upload,
            "ACCEPTED_FORMATS": ", ".join(f".{ext}" for ext in settings.allowed_extensions),
            "IS_ADMIN_AUTHENTICATED": bool(session.get("admin_authenticated")) if has_request_context() else False,
            "PLAN_LABELS": PLAN_LABELS,
            "HEALTH_TEMPLATE_AVAILABLE": bool(settings.health_template_path and settings.health_template_path.exists()),
            "DENTAL_TEMPLATE_AVAILABLE": bool(settings.dental_template_path and settings.dental_template_path.exists()),
            "csrf_token": get_csrf_token,
        }

    @app.errorhandler(RequestEntityTooLarge)
    def handle_too_large(_: RequestEntityTooLarge):
        return (
            render_template(
                "error.html",
                title="Arquivo muito grande",
                message=f"O limite de upload e de {settings.max_content_length_mb} MB por requisicao.",
            ),
            413,
        )

    @app.errorhandler(400)
    def handle_bad_request(_: Exception):
        return (
            render_template(
                "error.html",
                title="Requisicao invalida",
                message="Nao foi possivel concluir a operacao. Atualize a pagina e tente novamente.",
            ),
            400,
        )

    app.register_blueprint(public_bp)
    app.register_blueprint(admin_bp)
    return app
