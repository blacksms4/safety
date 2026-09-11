"""안전작업허가서 앱의 외부 서비스 설정 처리."""

from pathlib import Path


FIREBASE_CREDENTIAL_KEYS = (
    "type",
    "project_id",
    "private_key_id",
    "private_key",
    "client_email",
    "client_id",
    "auth_uri",
    "token_uri",
    "auth_provider_x509_cert_url",
    "client_x509_cert_url",
)


def resolve_firebase_credentials(secrets, environ):
    """Firebase 인증값을 Secrets 우선, 환경변수 차선으로 읽는다."""
    try:
        firebase_secrets = secrets.get("firebase", {}) if secrets is not None else {}
    except Exception:
        firebase_secrets = {}

    credentials = {}
    for key in FIREBASE_CREDENTIAL_KEYS:
        try:
            secret_value = firebase_secrets.get(key)
        except Exception:
            secret_value = None
        credentials[key] = secret_value or environ.get(
            f"STREAMLIT_FIREBASE_{key.upper()}"
        )

    missing = [key for key, value in credentials.items() if not value]
    if missing:
        raise ValueError("Firebase 설정이 완료되지 않았습니다: " + ", ".join(missing))

    credentials["private_key"] = credentials["private_key"].replace("\\n", "\n")
    return credentials


def _has_any_firebase_setting(secrets, environ):
    try:
        firebase_secrets = secrets.get("firebase", {}) if secrets is not None else {}
    except Exception:
        firebase_secrets = {}

    for key in FIREBASE_CREDENTIAL_KEYS:
        try:
            if firebase_secrets.get(key):
                return True
        except Exception:
            pass
        if environ.get(f"STREAMLIT_FIREBASE_{key.upper()}"):
            return True
    return False


def resolve_firebase_credentials_source(secrets, environ, json_key_path):
    """Firebase 인증 소스를 고른다.

    Streamlit Cloud 배포에서는 Secrets/환경변수를 우선 사용하고, 로컬 JSON 파일은
    설정값이 전혀 없을 때만 개발용 fallback으로 쓴다.
    """
    try:
        return "mapping", resolve_firebase_credentials(secrets, environ)
    except ValueError:
        if _has_any_firebase_setting(secrets, environ):
            raise

    json_key_path = Path(json_key_path)
    if json_key_path.exists():
        return "file", str(json_key_path)

    return "mapping", resolve_firebase_credentials(secrets, environ)
