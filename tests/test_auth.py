"""Tests for auth.py"""

import pytest
import msal
from unittest.mock import MagicMock, patch

import msgraphclient.auth as auth_mod
from msgraphclient.messages import EN_MESSAGES, get_messages
from msgraphclient.settings import GRAPH_DEFAULTS, GraphSettings


@pytest.fixture(autouse=True)
def _default_graph_auth_mode(monkeypatch: pytest.MonkeyPatch) -> None:
    """Isolate tests from local .env by pinning default auth mode."""
    monkeypatch.setenv("GRAPH_AUTH_MODE", "client_credentials")


def test_graph_client_missing_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphClient raises EnvironmentError when vars are missing."""
    monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    with pytest.raises(EnvironmentError, match="AZURE_TENANT_ID"):
        auth_mod.GraphClient()


def test_graph_client_uses_msal_token(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphClient stores token string when MSAL succeeds."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "site-id")

    fake_result = {"access_token": "fake-token-abc"}
    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = fake_result

    with (
        patch.object(msal, "ConfidentialClientApplication", return_value=mock_app),
        patch.object(auth_mod.GraphClient, "_load_site_info"),
    ):
        client = auth_mod.GraphClient()

    assert client._token == "fake-token-abc"
    assert client.authenticator.token == "fake-token-abc"


def test_graph_client_msal_error(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphClient raises RuntimeError when MSAL returns an error."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")

    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = {
        "error": "invalid_client",
        "error_description": "bad credentials",
    }

    with patch.object(msal, "ConfidentialClientApplication", return_value=mock_app):
        with pytest.raises(RuntimeError, match="invalid_client"):
            auth_mod.GraphClient()


def test_graph_client_reuses_provided_authenticator(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test GraphClient reuses a provided GraphAuthenticator instance."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    mock_authenticator = MagicMock()
    mock_authenticator.token = "token-from-authenticator"
    mock_authenticator.sharepoint_site_id = ""
    mock_authenticator.auth_mode = "client_credentials"

    client = auth_mod.GraphClient(authenticator=mock_authenticator)

    assert client.authenticator is mock_authenticator
    assert client._token == "token-from-authenticator"


def test_graph_authenticator_stores_sharepoint_site_id() -> None:
    """Test GraphAuthenticator stores sharepoint_site_id when passed."""
    auth = auth_mod.GraphAuthenticator(
        token="pre-built-token", sharepoint_site_id="site-custom"
    )

    assert auth.sharepoint_site_id == "site-custom"
    assert auth.token == "pre-built-token"


def test_graph_authenticator_initialization_with_explicit_token() -> None:
    """Test GraphAuthenticator stores token directly when provided."""
    auth = auth_mod.GraphAuthenticator(
        token="token-xyz",
        sharepoint_site_id="site-token",
    )

    assert auth.token == "token-xyz"
    assert auth.sharepoint_site_id == "site-token"


def test_graph_client_uses_delegated_mode(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test GraphClient acquires token via delegated mode when requested."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("GRAPH_AUTH_MODE", "delegated")
    monkeypatch.setenv("GRAPH_DELEGATED_LOGIN_MODE", "interactive")
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "site-id")
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    fake_result = {"access_token": "delegated-token-abc"}
    mock_app = MagicMock()
    mock_app.acquire_token_interactive.return_value = fake_result

    with (
        patch.object(msal, "PublicClientApplication", return_value=mock_app),
        patch.object(auth_mod.GraphClient, "_load_site_info"),
    ):
        client = auth_mod.GraphClient(auth_mode="delegated")

    assert client._token == "delegated-token-abc"
    assert client.authenticator.token == "delegated-token-abc"
    assert client.authenticator.auth_mode == "delegated"


def test_graph_client_delegated_mode_requires_env(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Delegated mode should require tenant and client id."""
    monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
    monkeypatch.setenv("GRAPH_AUTH_MODE", "delegated")
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    with pytest.raises(EnvironmentError, match="AZURE_TENANT_ID"):
        auth_mod.GraphClient(auth_mode="delegated")


def test_graph_client_credentials_requires_secret(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """client_credentials mode requires AZURE_CLIENT_SECRET in addition to common vars."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    with pytest.raises(EnvironmentError, match="AZURE_CLIENT_SECRET"):
        auth_mod.GraphClient()


def test_graph_delegated_does_not_require_secret(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Delegated mode should NOT require AZURE_CLIENT_SECRET."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "site-id")
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    fake_result = {"access_token": "delegated-no-secret"}
    mock_app = MagicMock()
    mock_app.acquire_token_interactive.return_value = fake_result

    with (
        patch.object(msal, "PublicClientApplication", return_value=mock_app),
        patch.object(auth_mod.GraphClient, "_load_site_info"),
    ):
        client = auth_mod.GraphClient(auth_mode="delegated")

    assert client._token == "delegated-no-secret"


def test_graph_authenticator_explicit_credentials() -> None:
    """Test GraphAuthenticator with explicit credentials (client_credentials)."""
    fake_result = {"access_token": "app-only-token"}
    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = fake_result

    with patch.object(msal, "ConfidentialClientApplication", return_value=mock_app):
        auth = auth_mod.GraphAuthenticator(
            tenant_id="tenant-id",
            client_id="client-id",
            client_secret="client-secret",
            auth_mode="client_credentials",
        )

    assert auth.auth_mode == "client_credentials"
    assert auth.token == "app-only-token"


def test_graph_authenticator_delegated_device_code_mode() -> None:
    """Delegated device_code mode should use MSAL device flow helpers."""
    mock_app = MagicMock()
    mock_app.initiate_device_flow.return_value = {
        "user_code": "ABCDEF",
        "message": "Open browser and enter code",
    }
    mock_app.acquire_token_by_device_flow.return_value = {
        "access_token": "delegated-device-token"
    }

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        auth = auth_mod.GraphAuthenticator(
            tenant_id="tenant-id",
            client_id="client-id",
            auth_mode="delegated",
            delegated_login_mode="device_code",
        )

    assert auth.auth_mode == "delegated"
    assert auth.token == "delegated-device-token"
    mock_app.initiate_device_flow.assert_called_once()
    mock_app.acquire_token_by_device_flow.assert_called_once()


def test_graph_authenticator_delegated_interactive_timeout() -> None:
    """Delegated interactive mode should surface timeout errors from MSAL."""
    mock_app = MagicMock()
    mock_app.acquire_token_interactive.return_value = {
        "error": "timeout",
        "error_description": "Timed out waiting for browser login",
    }

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        with pytest.raises(
            RuntimeError, match="Failed to acquire delegated token: timeout"
        ):
            auth_mod.GraphAuthenticator(
                tenant_id="tenant-id",
                client_id="client-id",
                auth_mode="delegated",
                delegated_login_mode="interactive",
            )


def test_graph_authenticator_delegated_interactive_cancellation() -> None:
    """Delegated interactive mode should surface user cancellation errors."""
    mock_app = MagicMock()
    mock_app.acquire_token_interactive.return_value = {
        "error": "access_denied",
        "error_description": "User canceled the authentication flow",
    }

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        with pytest.raises(
            RuntimeError, match="Failed to acquire delegated token: access_denied"
        ):
            auth_mod.GraphAuthenticator(
                tenant_id="tenant-id",
                client_id="client-id",
                auth_mode="delegated",
                delegated_login_mode="interactive",
            )


def test_graph_authenticator_delegated_interactive_invalid_scope() -> None:
    """Delegated interactive mode should surface invalid scope errors."""
    mock_app = MagicMock()
    mock_app.acquire_token_interactive.return_value = {
        "error": "invalid_scope",
        "error_description": "The provided scope is not valid",
    }

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        with pytest.raises(
            RuntimeError, match="Failed to acquire delegated token: invalid_scope"
        ):
            auth_mod.GraphAuthenticator(
                tenant_id="tenant-id",
                client_id="client-id",
                auth_mode="delegated",
                delegated_login_mode="interactive",
                delegated_scopes=["invalid.scope"],
            )


def test_graph_authenticator_delegated_device_code_cancellation() -> None:
    """Delegated device_code mode should surface user cancellation errors."""
    mock_app = MagicMock()
    mock_app.initiate_device_flow.return_value = {
        "user_code": "ABCDEF",
        "message": "Open browser and enter code",
    }
    mock_app.acquire_token_by_device_flow.return_value = {
        "error": "authorization_declined",
        "error_description": "User declined authentication",
    }

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        with pytest.raises(
            RuntimeError,
            match="Failed to acquire delegated token: authorization_declined",
        ):
            auth_mod.GraphAuthenticator(
                tenant_id="tenant-id",
                client_id="client-id",
                auth_mode="delegated",
                delegated_login_mode="device_code",
            )


def test_graph_authenticator_delegated_device_code_timeout() -> None:
    """Delegated device_code mode should surface polling timeout errors."""
    mock_app = MagicMock()
    mock_app.initiate_device_flow.return_value = {
        "user_code": "ABCDEF",
        "message": "Open browser and enter code",
    }
    mock_app.acquire_token_by_device_flow.return_value = {
        "error": "authorization_pending_timeout",
        "error_description": "Timed out waiting for user to complete auth",
    }

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        with pytest.raises(
            RuntimeError,
            match="Failed to acquire delegated token: authorization_pending_timeout",
        ):
            auth_mod.GraphAuthenticator(
                tenant_id="tenant-id",
                client_id="client-id",
                auth_mode="delegated",
                delegated_login_mode="device_code",
            )


def test_graph_authenticator_delegated_device_code_invalid_scope() -> None:
    """Delegated device_code mode should surface invalid scope at flow start."""
    mock_app = MagicMock()
    mock_app.initiate_device_flow.return_value = {
        "error": "invalid_scope",
        "error_description": "The provided scope is not valid",
    }

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        with pytest.raises(
            RuntimeError, match="Failed to acquire delegated token: invalid_scope"
        ):
            auth_mod.GraphAuthenticator(
                tenant_id="tenant-id",
                client_id="client-id",
                auth_mode="delegated",
                delegated_login_mode="device_code",
                delegated_scopes=["bad.scope"],
            )


def test_graph_settings_resolves_defaults_and_overrides(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """GraphSettings should centralize defaults and environment overrides."""
    monkeypatch.setenv("GRAPH_AUTH_MODE", "delegated")
    monkeypatch.setenv("AZURE_REDIRECT_URI", "http://localhost:8000")
    monkeypatch.setenv("GRAPH_DELEGATED_LOGIN_MODE", "device_code")
    monkeypatch.setenv("GRAPH_DELEGATED_SCOPES", "scope.one, scope.two openid profile")
    monkeypatch.setenv("GRAPH_AUTH_POPUP_SIZE", "640x720")

    settings = GraphSettings.from_sources(
        tenant_id="tenant-id",
        client_id="client-id",
        sharepoint_site_id="site-id",
    )

    assert settings.auth_mode == "delegated"
    assert settings.redirect_uri == "http://localhost:8000"
    assert settings.delegated_login_mode == "device_code"
    assert settings.delegated_scopes == ("scope.one", "scope.two")
    assert settings.auth_popup_size == "640x720"
    assert (
        settings.auth_mode == GRAPH_DEFAULTS.auth_mode
        or settings.auth_mode == "delegated"
    )


def test_get_messages_defaults_to_english() -> None:
    """Localized message resolution should default to the English bundle."""
    messages = get_messages()

    assert messages == EN_MESSAGES
    assert (
        messages.auth_success_html_text
        == "Authentication completed. You may close this window"
    )


def test_get_messages_loads_portuguese_bundle() -> None:
    """Portuguese bundle should be returned when pt locale is requested."""
    messages = get_messages("pt")

    assert (
        messages.auth_success_html_text
        == "Autenticacao concluida. Voce pode fechar esta janela"
    )
    assert (
        messages.invalid_msal_response
        == "Falha ao obter token: resposta invalida do MSAL"
    )


def test_get_messages_normalizes_regional_locale() -> None:
    """Regional locale tags should map to the base language bundle."""
    messages = get_messages("pt-BR")

    assert messages.missing_credentials_header.startswith(
        "Valores de configuracao ausentes"
    )


def test_get_messages_unknown_locale_falls_back_to_english() -> None:
    """Unsupported locales should fall back to the English message bundle."""
    messages = get_messages("fr")

    assert messages == EN_MESSAGES
