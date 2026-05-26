"""Tests for GraphAuthenticator site-discovery methods."""

from unittest.mock import MagicMock

import pytest

from msgraphtest.auth import GraphAuthenticator


@pytest.fixture()
def env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Set up environment variables required for GraphAuthenticator initialization."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")


def _mock_client() -> MagicMock:
    """Create a mock GraphClient-like object for authenticator tests."""
    client = MagicMock()
    client.get.side_effect = [
        {
            "id": "site-1",
            "name": "MySite",
            "displayName": "My Site",
            "webUrl": "https://contoso/sites/site-1",
        },
        {"value": [{"id": "drive-1", "name": "Documents"}]},
        {"value": [{"id": "list-1", "displayName": "Tasks"}]},
    ]
    return client


def test_get_site_contents_combines_site_drives_lists(env: None) -> None:
    """Test that get_site_contents returns site metadata and resource lists."""
    mock_client = _mock_client()

    auth = GraphAuthenticator(sharepoint_site_id="site-1", client=mock_client)
    result = auth.get_site_contents()

    assert result["site"]["id"] == "site-1"
    assert result["drives"][0]["id"] == "drive-1"
    assert result["lists"][0]["id"] == "list-1"


def test_get_site_contents_missing_site_id(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphAuthenticator raises if SHAREPOINT_SITE_ID is missing."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    with pytest.raises(EnvironmentError, match="SHAREPOINT_SITE_ID"):
        GraphAuthenticator(client=MagicMock())
