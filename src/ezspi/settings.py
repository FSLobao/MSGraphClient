"""Shared configuration defaults and resolution helpers for ezspi."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
import os

from ezspi.messages import get_messages


@dataclass(frozen=True, slots=True)
class Defaults:
    """Centralized defaults used by the authentication and client layers."""

    graph_api_base_url: str = "https://graph.microsoft.com/v1.0"
    graph_scopes: tuple[str, ...] = ("https://graph.microsoft.com/.default",)
    auth_mode: str = "client_credentials"
    redirect_uri: str = "http://localhost"
    delegated_login_mode: str = "interactive"
    delegated_scopes: tuple[str, ...] = ("https://graph.microsoft.com/Sites.Selected",)
    auth_popup_size: str = "520x680"


DEFAULTS = Defaults()

_MSAL_RESERVED_SCOPES: frozenset[str] = frozenset(
    {"openid", "profile", "offline_access"}
)


@dataclass(frozen=True, slots=True)
class Settings:
    """Resolved runtime settings for Graph authentication and SharePoint access."""

    tenant_id: str
    client_id: str
    client_secret: str
    sharepoint_site_id: str
    auth_mode: str
    redirect_uri: str
    delegated_scopes: tuple[str, ...]
    delegated_login_mode: str
    auth_popup_size: str

    @classmethod
    def from_sources(
        cls,
        *,
        tenant_id: str = "",
        client_id: str = "",
        client_secret: str = "",
        sharepoint_site_id: str = "",
        auth_mode: str | None = None,
        redirect_uri: str | None = None,
        delegated_scopes: Sequence[str] | None = None,
        delegated_scopes_raw: str | None = None,
        delegated_login_mode: str | None = None,
        auth_popup_size: str | None = None,
        env: Mapping[str, str] | None = None,
        defaults: Defaults = DEFAULTS,
    ) -> "Settings":
        """Resolve configuration values using explicit arguments, environment, then defaults."""

        env_mapping = env or os.environ

        resolved_auth_mode = _normalize_choice(
            auth_mode if auth_mode is not None else env_mapping.get("GRAPH_AUTH_MODE"),
            defaults.auth_mode,
        )
        resolved_redirect_uri = (
            redirect_uri
            or env_mapping.get("AZURE_REDIRECT_URI")
            or defaults.redirect_uri
        )
        resolved_login_mode = _normalize_choice(
            delegated_login_mode
            if delegated_login_mode is not None
            else env_mapping.get("GRAPH_DELEGATED_LOGIN_MODE"),
            defaults.delegated_login_mode,
        )
        resolved_popup_size = (
            auth_popup_size
            or env_mapping.get("GRAPH_AUTH_POPUP_SIZE")
            or defaults.auth_popup_size
        )

        if resolved_auth_mode not in ("client_credentials", "delegated"):
            raise ValueError(get_messages().unsupported_graph_auth_mode)
        if resolved_login_mode not in ("interactive", "device_code"):
            raise ValueError(get_messages().invalid_delegated_login_mode)

        if delegated_scopes is not None:
            resolved_scopes = tuple(_dedupe_scopes(delegated_scopes))
        else:
            raw_scopes = (
                delegated_scopes_raw
                if delegated_scopes_raw is not None
                else env_mapping.get("GRAPH_DELEGATED_SCOPES", "")
            )
            resolved_scopes = _parse_delegated_scopes(raw_scopes, defaults)

        return cls(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            sharepoint_site_id=sharepoint_site_id,
            auth_mode=resolved_auth_mode,
            redirect_uri=resolved_redirect_uri,
            delegated_scopes=resolved_scopes,
            delegated_login_mode=resolved_login_mode,
            auth_popup_size=resolved_popup_size,
        )


def _normalize_choice(value: str | None, fallback: str) -> str:
    """Normalize an option string using a fallback when no value is provided."""

    resolved = value if value is not None else fallback
    return resolved.strip().lower().replace("-", "_")


def _dedupe_scopes(scopes: Sequence[str]) -> list[str]:
    """Return scopes in order while removing duplicates and reserved OIDC entries."""

    deduped: list[str] = []
    for scope in scopes:
        normalized = scope.strip()
        if (
            normalized
            and normalized not in deduped
            and normalized not in _MSAL_RESERVED_SCOPES
        ):
            deduped.append(normalized)
    return deduped


def _parse_delegated_scopes(
    raw_scopes: str,
    defaults: Defaults = DEFAULTS,
) -> tuple[str, ...]:
    """Parse delegated scopes from a delimited string or fall back to defaults."""

    if not raw_scopes.strip():
        return defaults.delegated_scopes

    scopes = _dedupe_scopes(raw_scopes.replace(",", " ").split())
    return tuple(scopes) if scopes else defaults.delegated_scopes


def parse_popup_size(
    raw_size: str | None, defaults: Defaults = DEFAULTS
) -> tuple[int, int]:
    """Parse a WIDTHxHEIGHT popup size string, falling back to the default value."""

    value = (raw_size or defaults.auth_popup_size).lower().replace(",", "x")
    try:
        width_text, height_text = value.split("x", 1)
        return int(width_text), int(height_text)
    except (ValueError, TypeError):
        fallback_width, fallback_height = defaults.auth_popup_size.lower().split("x", 1)
        return int(fallback_width), int(fallback_height)

