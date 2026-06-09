"""Tests for ezspi.settings."""

from ezspi.settings import DEFAULTS, Settings, parse_popup_size


def test_parse_popup_size_uses_default_on_invalid_input() -> None:
    """Invalid popup size strings should fall back to the centralized default."""

    assert parse_popup_size("not-a-size") == (520, 680)
    assert parse_popup_size(None) == (520, 680)
    assert DEFAULTS.auth_popup_size == "520x680"


def test_graph_settings_preserves_explicit_delegated_scopes() -> None:
    """Explicit delegated scopes should take precedence over environment defaults."""

    settings = Settings.from_sources(
        tenant_id="tenant-id",
        client_id="client-id",
        sharepoint_site_id="site-id",
        delegated_scopes=["scope.a", "openid", "scope.b", "scope.a"],
    )

    assert settings.delegated_scopes == ("scope.a", "scope.b")

