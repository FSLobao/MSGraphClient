"""Tests for lists.py"""

import pytest
from unittest.mock import MagicMock, patch

import msgraphtest.lists as lists_mod


@pytest.fixture()
def env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Set up environment variables required for SharePoint list operations."""
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "site-xyz")
    monkeypatch.setenv("SHAREPOINT_LIST_ID", "list-abc")
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")


def _mock_client(return_value: dict | None = None) -> MagicMock:
    """Create a mock GraphClient instance for testing.

    Args:
        return_value: Return value for get() calls. Defaults to empty dict.

    Returns:
        A MagicMock object configured to simulate GraphClient behavior.
    """
    client = MagicMock()
    client.get.return_value = return_value or {}
    client.post.return_value = {"id": "42", "fields": {"Title": "New Item"}}
    client.patch.return_value = {"Title": "Updated Item"}
    return client


def test_get_list_items_returns_value(env: None) -> None:
    """Test that get_list_items returns the value array from API response."""
    items = [
        {"id": "1", "fields": {"Title": "Item A"}},
        {"id": "2", "fields": {"Title": "Item B"}},
    ]
    mock_client = _mock_client(return_value={"value": items})

    with patch.object(lists_mod, "GraphClient", return_value=mock_client):
        result = lists_mod.get_list_items()

    assert result == items
    mock_client.get.assert_called_once()


def test_get_list_items_missing_site_id(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that get_list_items raises EnvironmentError when SHAREPOINT_SITE_ID is missing."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    with patch.object(lists_mod, "GraphClient"):
        with pytest.raises(EnvironmentError, match="SHAREPOINT_SITE_ID"):
            lists_mod.get_list_items()


def test_get_list_items_missing_list_id(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that get_list_items raises EnvironmentError when SHAREPOINT_LIST_ID is missing."""
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "site-xyz")
    monkeypatch.delenv("SHAREPOINT_LIST_ID", raising=False)

    with patch.object(lists_mod, "GraphClient"):
        with pytest.raises(EnvironmentError, match="SHAREPOINT_LIST_ID"):
            lists_mod.get_list_items()


def test_create_list_item(env: None) -> None:
    """Test that create_list_item sends field data and returns the created item."""
    mock_client = _mock_client()

    with patch.object(lists_mod, "GraphClient", return_value=mock_client):
        result = lists_mod.create_list_item({"Title": "New Item", "Status": "Active"})

    mock_client.post.assert_called_once()
    assert result["id"] == "42"


def test_update_list_item(env: None) -> None:
    """Test that update_list_item sends updated fields to the Graph API."""
    mock_client = _mock_client()

    with patch.object(lists_mod, "GraphClient", return_value=mock_client):
        result = lists_mod.update_list_item("42", {"Status": "Closed"})

    mock_client.patch.assert_called_once()
    assert result["Title"] == "Updated Item"


def test_get_list_columns_returns_value(env: None) -> None:
    """Test that get_list_columns returns the value array from the columns endpoint."""
    columns = [
        {"name": "Title", "displayName": "Title"},
        {"name": "field_1", "displayName": "Customer Name"},
    ]
    mock_client = _mock_client(return_value={"value": columns})

    with patch.object(lists_mod, "GraphClient", return_value=mock_client):
        result = lists_mod.get_list_columns()

    assert result == columns
    mock_client.get.assert_called_once()


def test_get_list_views_returns_value(env: None) -> None:
    """Test that get_list_views returns the value array from the views endpoint."""
    views = [
        {"id": "view-1", "name": "All Items"},
        {"id": "view-2", "name": "Active Only"},
    ]
    mock_client = _mock_client(return_value={"value": views})

    with patch.object(lists_mod, "GraphClient", return_value=mock_client):
        result = lists_mod.get_list_views()

    assert result == views
    mock_client.get.assert_called_once()


def test_get_list_view_columns_returns_value(env: None) -> None:
    """Test that get_list_view_columns returns the value array for a specific view."""
    columns = [
        {"name": "Title", "displayName": "Title"},
        {"name": "field_2", "displayName": "Status"},
    ]
    mock_client = _mock_client(return_value={"value": columns})

    with patch.object(lists_mod, "GraphClient", return_value=mock_client):
        result = lists_mod.get_list_view_columns("view-1")

    assert result == columns
    call_path = mock_client.get.call_args[0][0]
    assert "views/view-1/columns" in call_path
