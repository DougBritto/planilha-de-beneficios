from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _optional_path(value: str | None, default: Path | None = None) -> Path | None:
    candidate = value.strip() if value else ""
    if candidate:
        return Path(candidate).expanduser()
    return default


def _to_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "t", "yes", "y", "on"}


@dataclass(frozen=True)
class AppSettings:
    base_dir: Path
    data_dir: Path
    upload_dir: Path
    output_dir: Path
    db_path: Path
    app_title: str
    app_subtitle: str
    admin_password: str
    admin_password_hash: str
    secret_key: str
    data_encryption_key: str
    default_sheet_name: str
    include_source_columns: bool
    max_content_length_mb: int
    max_files_per_upload: int
    max_name_length: int
    max_email_length: int
    max_note_length: int
    session_cookie_secure: bool
    session_lifetime_minutes: int
    admin_session_idle_minutes: int
    login_max_attempts: int
    login_lock_minutes: int
    upload_rate_limit_count: int
    upload_rate_limit_window_minutes: int
    block_duplicate_files: bool
    allowed_extensions: tuple[str, ...]
    health_template_path: Path | None
    dental_template_path: Path | None
    port: int
    debug: bool

    @property
    def max_content_length_bytes(self) -> int:
        return self.max_content_length_mb * 1024 * 1024

    def base_template_for(self, plan_type: str) -> Path | None:
        if plan_type == "saude":
            return self.health_template_path
        if plan_type == "odonto":
            return self.dental_template_path
        return None

    @classmethod
    def from_env(cls) -> "AppSettings":
        base_dir = Path(__file__).resolve().parent.parent
        data_dir = Path(os.getenv("DATA_DIR", base_dir / "data"))
        default_health_template = base_dir / "PLANILHA BASE UNIMED PLANO DE SA\u00daDE..xlsx"
        default_dental_template = base_dir / "PLANILHA BASE UNIMED PLANO ODONTOL\u00d3GICO.xlsx"
        return cls(
            base_dir=base_dir,
            data_dir=data_dir,
            upload_dir=data_dir / "uploads",
            output_dir=data_dir / "outputs",
            db_path=data_dir / "app.db",
            app_title=os.getenv("APP_TITLE", "Consolidador de Planilhas"),
            app_subtitle=os.getenv(
                "APP_SUBTITLE",
                "Receba planilhas por um link unico e gere um consolidado rastreavel.",
            ),
            admin_password=os.getenv("ADMIN_PASSWORD", "troque-esta-senha"),
            admin_password_hash=os.getenv("ADMIN_PASSWORD_HASH", "").strip(),
            secret_key=os.getenv("SECRET_KEY", "troque-esta-secret-key"),
            data_encryption_key=os.getenv("DATA_ENCRYPTION_KEY", "").strip(),
            default_sheet_name=os.getenv("SHEET_NAME", "").strip(),
            include_source_columns=_to_bool(os.getenv("INCLUDE_SOURCE_COLUMNS"), True),
            max_content_length_mb=max(1, int(os.getenv("MAX_CONTENT_LENGTH_MB", "100"))),
            max_files_per_upload=max(1, int(os.getenv("MAX_FILES_PER_UPLOAD", "20"))),
            max_name_length=max(20, int(os.getenv("MAX_NAME_LENGTH", "120"))),
            max_email_length=max(20, int(os.getenv("MAX_EMAIL_LENGTH", "180"))),
            max_note_length=max(50, int(os.getenv("MAX_NOTE_LENGTH", "1000"))),
            session_cookie_secure=_to_bool(os.getenv("SESSION_COOKIE_SECURE"), False),
            session_lifetime_minutes=max(15, int(os.getenv("SESSION_LIFETIME_MINUTES", "480"))),
            admin_session_idle_minutes=max(5, int(os.getenv("ADMIN_SESSION_IDLE_MINUTES", "60"))),
            login_max_attempts=max(3, int(os.getenv("LOGIN_MAX_ATTEMPTS", "6"))),
            login_lock_minutes=max(1, int(os.getenv("LOGIN_LOCK_MINUTES", "15"))),
            upload_rate_limit_count=max(1, int(os.getenv("UPLOAD_RATE_LIMIT_COUNT", "12"))),
            upload_rate_limit_window_minutes=max(1, int(os.getenv("UPLOAD_RATE_LIMIT_WINDOW_MINUTES", "15"))),
            block_duplicate_files=_to_bool(os.getenv("BLOCK_DUPLICATE_FILES"), False),
            allowed_extensions=("xlsx", "xls", "xlsm", "csv"),
            health_template_path=_optional_path(
                os.getenv("HEALTH_TEMPLATE_PATH"),
                default_health_template if default_health_template.exists() else None,
            ),
            dental_template_path=_optional_path(
                os.getenv("DENTAL_TEMPLATE_PATH"),
                default_dental_template if default_dental_template.exists() else None,
            ),
            port=int(os.getenv("PORT", "5000")),
            debug=_to_bool(os.getenv("FLASK_DEBUG"), False),
        )
