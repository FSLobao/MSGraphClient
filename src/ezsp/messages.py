"""Localized user-facing messages for ezsp."""

from __future__ import annotations

from dataclasses import dataclass
from dataclasses import fields
from functools import lru_cache
import json
import os
from pathlib import Path


@dataclass(frozen=True, slots=True)
class Messages:
    """Container for translatable user-facing text."""

    missing_credentials_header: str
    missing_credentials_instructions: str
    missing_credentials_app_only: str
    missing_credentials_delegated: str
    missing_credentials_footer: str
    invalid_delegated_login_mode: str
    unsupported_graph_auth_mode: str
    invalid_msal_response: str
    invalid_delegated_msal_response: str
    token_acquire_failed: str
    delegated_token_acquire_failed: str
    device_code_prompt: str
    auth_success_html_text: str
    no_valid_token: str


_DEFAULT_LOCALE = "en"
_LOCALE_DIR = Path(__file__).with_name("locales")
_MESSAGE_KEYS = tuple(field.name for field in fields(Messages))
_ENGLISH_FALLBACK: dict[str, str] = {
    "missing_credentials_header": "Missing configuration values (argument or environment): {missing}",
    "missing_credentials_instructions": "Create a .env file in your project root:\n  {env_path}",
    "missing_credentials_app_only": (
        "Minimum example for client_credentials authentication:\n\n"
        "  AZURE_TENANT_ID=your-tenant-id\n"
        "  AZURE_CLIENT_ID=your-client-id\n"
        "  AZURE_CLIENT_SECRET=your-client-secret\n"
        "  SHAREPOINT_SITE_ID=your-site-id"
    ),
    "missing_credentials_delegated": "For delegated authentication, AZURE_CLIENT_SECRET is not required.",
    "missing_credentials_footer": "Docs: {docs_url}\n  Repo: {repo_url}",
    "invalid_delegated_login_mode": "GRAPH_DELEGATED_LOGIN_MODE must be 'interactive' or 'device_code'",
    "unsupported_graph_auth_mode": "Unsupported GRAPH_AUTH_MODE. Use 'client_credentials' or 'delegated'.",
    "invalid_msal_response": "Failed to acquire token: invalid response from MSAL",
    "invalid_delegated_msal_response": "Failed to acquire delegated token: invalid response from MSAL",
    "token_acquire_failed": "Failed to acquire token: {error} - {description}",
    "delegated_token_acquire_failed": "Failed to acquire delegated token: {error} - {description}",
    "device_code_prompt": "Complete device authentication to continue.",
    "auth_success_html_text": "Authentication completed. This window should close shortly, if not, please close it manually.",
    "no_valid_token": "No valid token available. Ensure credentials are configured so Authenticator can acquire a token.",
}


def _read_locale_bundle(locale: str) -> dict[str, str]:
    """Read a locale bundle from JSON, returning an empty mapping on failure."""
    file_path = _LOCALE_DIR / f"{locale}.json"
    if not file_path.is_file():
        return {}

    try:
        with open(file_path, "r", encoding="utf-8") as file_obj:
            payload = json.load(file_obj)
    except (OSError, ValueError):
        return {}

    if not isinstance(payload, dict):
        return {}

    return {
        key: value
        for key, value in payload.items()
        if isinstance(key, str) and isinstance(value, str)
    }


def _build_messages(bundle: dict[str, str], fallback: dict[str, str]) -> Messages:
    """Build a complete Messages instance using fallback values when needed."""
    values: dict[str, str] = {}
    for key in _MESSAGE_KEYS:
        resolved = bundle.get(key)
        values[key] = (
            resolved if isinstance(resolved, str) and resolved else fallback[key]
        )
    return Messages(**values)


@lru_cache(maxsize=1)
def _load_catalog() -> dict[str, Messages]:
    """Load the message catalog from locale bundles with English fallback."""
    english_bundle = {**_ENGLISH_FALLBACK, **_read_locale_bundle("en")}
    english_messages = _build_messages(english_bundle, _ENGLISH_FALLBACK)
    portuguese_messages = _build_messages(_read_locale_bundle("pt"), english_bundle)
    return {
        "en": english_messages,
        "pt": portuguese_messages,
    }


EN_MESSAGES = _load_catalog()["en"]
MESSAGE_CATALOG: dict[str, Messages] = _load_catalog()


def _normalize_locale(locale: str | None) -> str:
    """Normalize locale input and environment value to a base language code."""
    resolved_locale = locale or os.environ.get("GRAPH_LOCALE") or _DEFAULT_LOCALE
    normalized = resolved_locale.strip().lower().replace("_", "-")
    return normalized.split("-", 1)[0]


def get_messages(locale: str | None = None) -> Messages:
    """Return the message bundle for a requested locale, defaulting to English."""
    language = _normalize_locale(locale)
    return MESSAGE_CATALOG.get(language, EN_MESSAGES)

