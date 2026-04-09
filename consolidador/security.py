from __future__ import annotations

import secrets
from datetime import datetime, timedelta
from functools import wraps

from flask import abort, current_app, flash, redirect, request, session, url_for

from .config import AppSettings


CSRF_SESSION_KEY = "_csrf_token"


def get_csrf_token() -> str:
    token = session.get(CSRF_SESSION_KEY)
    if not token:
        token = secrets.token_urlsafe(32)
        session[CSRF_SESSION_KEY] = token
    return token


def _validate_csrf_token() -> None:
    if request.method not in {"POST", "PUT", "PATCH", "DELETE"}:
        return
    token = request.form.get("csrf_token") or request.headers.get("X-CSRFToken")
    import hmac

    session_token = session.get(CSRF_SESSION_KEY, "")
    if not token or not session_token or not hmac.compare_digest(token, session_token):
        abort(400)


def _is_admin_session_expired() -> bool:
    settings: AppSettings = current_app.config["APP_SETTINGS"]
    last_seen = session.get("admin_last_seen")
    if not last_seen:
        return True
    try:
        seen_at = datetime.fromisoformat(last_seen)
    except ValueError:
        return True
    return datetime.now() - seen_at > timedelta(minutes=settings.admin_session_idle_minutes)


def admin_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if not session.get("admin_authenticated"):
            return redirect(url_for("admin.login"))
        if _is_admin_session_expired():
            session.clear()
            flash("Sua sessao administrativa expirou por inatividade. Entre novamente.", "warning")
            return redirect(url_for("admin.login"))
        session["admin_last_seen"] = datetime.now().isoformat(timespec="seconds")
        return view(*args, **kwargs)

    return wrapped_view


def init_security(app) -> None:
    @app.before_request
    def apply_session_defaults():
        session.permanent = True

    @app.before_request
    def csrf_protect():
        if request.endpoint == "static":
            return None
        return _validate_csrf_token()

    @app.after_request
    def set_security_headers(response):
        response.headers.setdefault("X-Frame-Options", "DENY")
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("Referrer-Policy", "same-origin")
        response.headers.setdefault("Cross-Origin-Opener-Policy", "same-origin")
        response.headers.setdefault("Cross-Origin-Resource-Policy", "same-origin")
        response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
        response.headers.setdefault(
            "Content-Security-Policy",
            "default-src 'self'; "
            "style-src 'self'; "
            "script-src 'self'; "
            "img-src 'self' data:; "
            "font-src 'self' data:; "
            "form-action 'self'; "
            "base-uri 'self'; "
            "frame-ancestors 'none'",
        )
        if request.path.startswith("/admin"):
            response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, private, max-age=0"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
            response.headers["X-Robots-Tag"] = "noindex, nofollow, noarchive"
        return response


def is_admin_password_valid(password: str) -> bool:
    import hmac
    from werkzeug.security import check_password_hash

    settings: AppSettings = current_app.config["APP_SETTINGS"]
    if settings.admin_password_hash:
        return check_password_hash(settings.admin_password_hash, password)
    return hmac.compare_digest(password, settings.admin_password)
