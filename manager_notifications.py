"""관리자 명단과 안전작업허가서 이메일 알림 기능."""

import re
import smtplib
import ssl
from email.message import EmailMessage


_EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


def validate_manager(
    name, email, team_leader_name=None, team_leader_department=None
):
    """관리자 입력값을 검증하고 저장 가능한 형태로 정규화한다."""
    normalized_name = str(name or "").strip()
    normalized_email = str(email or "").strip().lower()
    normalized_team_leader_name = str(team_leader_name or "").strip()
    normalized_team_leader_department = str(
        team_leader_department or ""
    ).strip()

    if not normalized_name:
        raise ValueError("관리자 이름을 입력해 주세요.")
    if not _EMAIL_PATTERN.fullmatch(normalized_email):
        raise ValueError("올바른 메일주소를 입력해 주세요.")
    if team_leader_name is not None or team_leader_department is not None:
        if not normalized_team_leader_name:
            raise ValueError("승인자(팀장) 이름을 입력해 주세요.")
        if not normalized_team_leader_department:
            raise ValueError("승인자(팀장) 부서를 입력해 주세요.")

    manager = {"name": normalized_name, "email": normalized_email}
    if team_leader_name is not None:
        manager["team_leader_name"] = normalized_team_leader_name
    if team_leader_department is not None:
        manager["team_leader_department"] = normalized_team_leader_department
    return manager


def attach_manager_notification(form_data, manager, enabled):
    """제출 데이터에 선택 당시의 관리자 정보와 발송 상태를 복사한다."""
    result = dict(form_data)
    result.update(
        {
            "manager_id": manager["id"],
            "manager_name": manager["name"],
            "manager_email": manager["email"],
            "team_leader_name": manager.get("team_leader_name", ""),
            "team_leader_department": manager.get(
                "team_leader_department", ""
            ),
            "email_notification_requested": bool(enabled),
            "email_status": "pending" if enabled else "disabled",
        }
    )
    return result


def form_instance_key(form_data, fallback_index):
    """제출 건별 Streamlit 위젯에 사용할 충돌 없는 키를 반환한다."""
    return str(
        form_data.get("_doc_id")
        or f"{form_data.get('id', 'legacy')}_{fallback_index}"
    )


def _as_bool(value, default=True):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def resolve_smtp_config(secrets, environ):
    """Streamlit Secrets 또는 환경변수에서 SMTP 설정을 읽는다."""
    try:
        email_secrets = secrets.get("email", {}) if secrets is not None else {}
    except Exception:
        email_secrets = {}

    def setting(secret_key, environment_key, default=None):
        try:
            secret_value = email_secrets.get(secret_key)
        except Exception:
            secret_value = None
        return secret_value if secret_value not in (None, "") else environ.get(
            environment_key, default
        )

    config = {
        "host": setting("smtp_host", "SMTP_HOST"),
        "port": int(setting("smtp_port", "SMTP_PORT", 587)),
        "username": setting("smtp_username", "SMTP_USERNAME", ""),
        "password": setting("smtp_password", "SMTP_PASSWORD", ""),
        "sender_email": setting("sender_email", "SMTP_SENDER_EMAIL"),
        "use_tls": _as_bool(setting("use_tls", "SMTP_USE_TLS", True)),
    }
    missing = [key for key in ("host", "sender_email") if not config[key]]
    if missing:
        raise ValueError("SMTP 설정이 완료되지 않았습니다: " + ", ".join(missing))
    return config


def build_submission_message(form_data, sender_email):
    """안전작업허가서 제출 요약 메일을 만든다."""
    recipient = validate_manager(
        form_data.get("manager_name"), form_data.get("manager_email")
    )
    location = form_data.get("work_location") or "작업장소 미입력"
    message = EmailMessage()
    message["Subject"] = f"[안전작업허가서] {location} 작업 신청"
    message["From"] = sender_email
    message["To"] = recipient["email"]
    message.set_content(
        "\n".join(
            [
                f"{recipient['name']} 관리자님, 안전작업허가서가 제출되었습니다.",
                "",
                f"허가서 ID: {form_data.get('id', '')}",
                f"작업일자: {form_data.get('work_date', '')}",
                f"작업장소: {location}",
                f"작업종류: {form_data.get('work_type', '')}",
                f"업체명: {form_data.get('company_name', '')}",
                f"작업자: {form_data.get('worker_name', '')}",
                f"작업인원: {form_data.get('worker_count') or ''}",
                f"작업내용: {form_data.get('work_description', '')}",
                "",
                "관리자 화면에서 제출 내용을 확인해 주세요.",
            ]
        )
    )
    return message


def _attach_files(message, attachments):
    for attachment in attachments or []:
        content = attachment["content"]
        if hasattr(content, "getvalue"):
            content = content.getvalue()
        maintype, subtype = attachment["mime_type"].split("/", 1)
        message.add_attachment(
            content,
            maintype=maintype,
            subtype=subtype,
            filename=attachment["filename"],
        )


def send_submission_email(
    form_data, smtp_config, attachments=None, smtp_factory=smtplib.SMTP
):
    """선택된 관리자에게 SMTP로 제출 요약 메일을 발송한다."""
    message = build_submission_message(form_data, smtp_config["sender_email"])
    _attach_files(message, attachments)
    with smtp_factory(
        smtp_config["host"], smtp_config["port"], timeout=30
    ) as smtp:
        smtp.ehlo()
        if smtp_config.get("use_tls", True):
            smtp.starttls(context=ssl.create_default_context())
            smtp.ehlo()
        if smtp_config.get("username"):
            smtp.login(smtp_config["username"], smtp_config.get("password", ""))
        smtp.send_message(message)


class ManagerRepository:
    """Firestore의 관리자 명단을 관리한다."""

    def __init__(self, database):
        self._collection = database.collection("managers")

    def list_all(self):
        managers = []
        for document in self._collection.stream():
            data = document.to_dict() or {}
            if data.get("name") and data.get("email"):
                manager = {
                    "id": document.id,
                    "name": data["name"],
                    "email": data["email"],
                }
                if "team_leader_name" in data:
                    manager["team_leader_name"] = data.get("team_leader_name", "")
                if "team_leader_department" in data:
                    manager["team_leader_department"] = data.get(
                        "team_leader_department", ""
                    )
                managers.append(manager)
        return sorted(managers, key=lambda manager: manager["name"].casefold())

    def create(
        self,
        name,
        email,
        team_leader_name=None,
        team_leader_department=None,
    ):
        manager = validate_manager(
            name, email, team_leader_name, team_leader_department
        )
        for existing_manager in self.list_all():
            if (
                existing_manager["name"].casefold() == manager["name"].casefold()
                and existing_manager["email"].casefold() == manager["email"].casefold()
            ):
                leader_updates = {
                    key: manager[key]
                    for key in (
                        "team_leader_name",
                        "team_leader_department",
                    )
                    if manager.get(key)
                }
                if leader_updates:
                    self._collection.document(existing_manager["id"]).update(
                        leader_updates
                    )
                return existing_manager["id"]
        _write_time, reference = self._collection.add(manager)
        return reference.id

    def update(
        self,
        manager_id,
        name,
        email,
        team_leader_name=None,
        team_leader_department=None,
    ):
        manager = validate_manager(
            name, email, team_leader_name, team_leader_department
        )
        self._collection.document(manager_id).update(manager)

    def delete(self, manager_id):
        self._collection.document(manager_id).delete()


class NotificationSettingsRepository:
    """Firestore에 자동 메일 발송 전역 설정을 저장한다."""

    def __init__(self, database):
        self._document = database.collection("app_settings").document(
            "email_notifications"
        )

    def is_enabled(self):
        snapshot = self._document.get()
        if not snapshot.exists:
            return False
        return bool((snapshot.to_dict() or {}).get("enabled", False))

    def set_enabled(self, enabled):
        self._document.set({"enabled": bool(enabled)}, merge=True)
