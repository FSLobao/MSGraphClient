"""Tests for site.py"""

import pytest
from unittest.mock import MagicMock, patch

import msgraphtest.site as site_mod


@pytest.fixture()
def env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Set up environment variables required for SharePoint site operations."""
    monkeypatch.setenv(
        "SHAREPOINT_SITE_ID", "contoso.sharepoint.com,site-guid,web-guid"
    )
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")


def _mock_client() -> MagicMock:
    """Create a mock GraphClient instance for site discovery tests."""
    client = MagicMock()
    client.get.side_effect = [
        {"id": "site-1", "name": "MySite", "displayName": "My Site"},
        {"value": [{"id": "drive-1", "name": "Documents"}]},
        {"value": [{"id": "list-1", "displayName": "Tasks"}]},
    ]
    return client


def test_get_site_contents_combines_site_drives_lists(env: None) -> None:
    """Test that get_site_contents returns site metadata and resource lists."""
    mock_client = _mock_client()

    with patch.object(site_mod, "GraphClient", return_value=mock_client):
        result = site_mod.get_site_contents()

    assert result["site"]["id"] == "site-1"
    assert result["drives"][0]["id"] == "drive-1"
    assert result["lists"][0]["id"] == "list-1"


def test_get_site_contents_missing_site_id(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that get_site_contents raises if SHAREPOINT_SITE_ID is missing."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    with patch.object(site_mod, "GraphClient"):
        with pytest.raises(EnvironmentError, match="SHAREPOINT_SITE_ID"):
            site_mod.get_site_contents()
