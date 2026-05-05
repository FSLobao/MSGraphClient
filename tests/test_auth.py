"""Tests for auth.py"""

import pytest
import msal
from unittest.mock import MagicMock, patch

import msgraphtest.auth as auth_mod


def test_get_access_token_missing_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that get_access_token raises EnvironmentError when vars are missing."""
    monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    with pytest.raises(EnvironmentError, match="AZURE_TENANT_ID"):
        auth_mod.get_access_token()


def test_get_access_token_success(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that get_access_token returns token string when MSAL succeeds."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")

    fake_result = {"access_token": "fake-token-abc"}
    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = fake_result

    with patch.object(msal, "ConfidentialClientApplication", return_value=mock_app):
        token = auth_mod.get_access_token()

    assert token == "fake-token-abc"


def test_get_access_token_msal_error(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that get_access_token raises RuntimeError when MSAL returns an error."""
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
            auth_mod.get_access_token()
