"""Tests for lists.py"""

from unittest.mock import MagicMock

import pytest
import requests

import python.lists as lists_mod


@pytest.fixture()
def env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Set up environment variables required for SharePoint list operations."""
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "site-xyz")
    monkeypatch.setenv("SHAREPOINT_LIST_ID", "list-abc")
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")


def _mock_client(return_value: dict | None = None) -> MagicMock:
    """Create a mock GraphClient instance for testing."""
    client = MagicMock()
    client.get.side_effect = [
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        return_value or {},
    ]
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

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_items()

    assert result == items
    assert mock_client.get.call_count == 2


def test_get_list_items_without_select_requests_all_fields(env: None) -> None:
    """Without select, get_list_items should not apply fields(select=...) filtering."""
    items = [{"id": "1", "fields": {"Title": "Item A", "field_1": "Alpha"}}]
    mock_client = _mock_client(return_value={"value": items})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_items(include_title=True, fields_only=False)

    assert result == items
    call_path = mock_client.get.call_args[0][0]
    assert "items?expand=fields" in call_path
    assert "fields(select=" not in call_path


def test_get_list_items_selects_specific_fields(env: None) -> None:
    """Test that get_list_items requests specific fields via expand=fields(select=...)."""
    items = [{"id": "1", "fields": {"Title": "Item A", "field_1": "Alpha"}}]
    mock_client = _mock_client(return_value={"value": items})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_items(select=["Title", "field_1"])

    assert result == items
    call_path = mock_client.get.call_args[0][0]
    assert "items?expand=fields(select=Title,field_1)" in call_path


def test_get_list_items_fields_only_returns_business_fields(env: None) -> None:
    """Test that get_list_items can return only the fields payload for each item."""
    items = [
        {"id": "1", "fields": {"Title": "Item A", "field_1": "Alpha"}},
        {"id": "2", "fields": {"Title": "Item B", "field_1": "Beta"}},
    ]
    mock_client = _mock_client(return_value={"value": items})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_items(select=["Title", "field_1"], fields_only=True)

    assert result == [
        {"Title": "Item A", "field_1": "Alpha"},
        {"Title": "Item B", "field_1": "Beta"},
    ]


def test_get_list_items_fields_only_without_select_filters_field_keys(
    env: None,
) -> None:
    """Without select, fields_only should keep only field_* keys and optional Title."""
    items = [
        {
            "id": "1",
            "fields": {
                "Title": "Item A",
                "field_1": "Alpha",
                "Author": "User A",
                "ContentType": "Item",
            },
        }
    ]

    mock_client_without_title = _mock_client(return_value={"value": items})
    list_client_without_title = lists_mod.GraphList(client=mock_client_without_title)
    result_without_title = list_client_without_title.get_list_items(fields_only=True)

    mock_client_with_title = _mock_client(return_value={"value": items})
    list_client_with_title = lists_mod.GraphList(client=mock_client_with_title)
    result_with_title = list_client_with_title.get_list_items(
        fields_only=True,
        include_title=True,
    )

    assert result_without_title == [{"field_1": "Alpha"}]
    assert result_with_title == [{"field_1": "Alpha", "Title": "Item A"}]


def test_get_list_items_can_include_title_without_fetching_all_fields(
    env: None,
) -> None:
    """Test that get_list_items can request Title plus explicit business fields."""
    items = [{"id": "1", "fields": {"Title": "Item A", "field_1": "Alpha"}}]
    mock_client = _mock_client(return_value={"value": items})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_items(
        select=["field_1"],
        include_title=True,
        fields_only=True,
    )

    assert result == [{"Title": "Item A", "field_1": "Alpha"}]
    call_path = mock_client.get.call_args[0][0]
    assert "items?expand=fields(select=Title,field_1)" in call_path


def test_get_list_items_fields_only_can_keep_item_id(env: None) -> None:
    """Test that fields_only mode can preserve the Graph list item id."""
    items = [{"id": "37", "fields": {"Title": "Item A", "field_1": "Alpha"}}]
    mock_client = _mock_client(return_value={"value": items})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_items(
        select=["field_1"],
        include_title=True,
        fields_only=True,
        include_item_id=True,
    )

    assert result == [{"Title": "Item A", "field_1": "Alpha", "id": "37"}]


def test_graph_list_initialization_loads_basic_metadata(env: None) -> None:
    """Test that GraphList validates access and stores basic list attributes."""
    mock_client = MagicMock()
    mock_client.get.return_value = {
        "id": "list-abc",
        "name": "MonitorRNI",
        "displayName": "Monitor RNI",
        "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
    }

    lst = lists_mod.GraphList(client=mock_client)

    assert lst.list_graph_id == "list-abc"
    assert lst.list_name == "MonitorRNI"
    assert lst.list_display_name == "Monitor RNI"
    assert lst.list_web_url.startswith("https://contoso.sharepoint.com")


def test_graph_list_initialization_with_explicit_arguments(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test GraphList accepts explicit list id and injected client."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)
    monkeypatch.delenv("SHAREPOINT_LIST_ID", raising=False)

    mock_client = MagicMock()
    mock_client.authenticator = MagicMock(sharepoint_site_id="site-custom")
    mock_client.get.return_value = {
        "id": "list-custom",
        "name": "CustomList",
        "displayName": "Custom List",
        "webUrl": "https://contoso.sharepoint.com/sites/custom/lists/customlist",
    }

    lst = lists_mod.GraphList(
        list_id="list-custom",
        client=mock_client,
    )

    assert lst.site_id == "site-custom"
    assert lst.list_id == "list-custom"
    assert lst.list_graph_id == "list-custom"
    mock_client.get.assert_called_once()
    assert "/sites/site-custom/lists/list-custom" in mock_client.get.call_args[0][0]


def test_graph_list_initialization_uses_site_id_from_client_authenticator(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test GraphList derives site_id from client.authenticator when omitted."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)
    monkeypatch.delenv("SHAREPOINT_LIST_ID", raising=False)

    mock_client = MagicMock()
    mock_client.authenticator = MagicMock(sharepoint_site_id="site-from-client")
    mock_client.get.return_value = {
        "id": "list-client",
        "name": "ClientList",
        "displayName": "Client List",
        "webUrl": "https://contoso.sharepoint.com/sites/client/lists/clientlist",
    }

    lst = lists_mod.GraphList(
        list_id="list-client",
        client=mock_client,
    )

    assert lst.site_id == "site-from-client"
    assert lst.list_id == "list-client"
    mock_client.get.assert_called_once()
    assert (
        "/sites/site-from-client/lists/list-client" in mock_client.get.call_args[0][0]
    )


def test_get_list_items_missing_site_id(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphList raises EnvironmentError when SHAREPOINT_SITE_ID is missing."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    with pytest.raises(EnvironmentError, match="SHAREPOINT_SITE_ID"):
        lists_mod.GraphList(client=MagicMock())


def test_get_list_items_missing_list_id(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphList raises EnvironmentError when SHAREPOINT_LIST_ID is missing."""
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "site-xyz")
    monkeypatch.delenv("SHAREPOINT_LIST_ID", raising=False)

    with pytest.raises(EnvironmentError, match="SHAREPOINT_LIST_ID"):
        lists_mod.GraphList(client=MagicMock())


def test_create_list_item(env: None) -> None:
    """Test that create_list_item sends field data and returns the created item."""
    mock_client = _mock_client()
    list_client = lists_mod.GraphList(client=mock_client)

    result = list_client.create_list_item({"Title": "New Item", "Status": "Active"})

    mock_client.post.assert_called_once()
    assert result["id"] == "42"


def test_update_list_item(env: None) -> None:
    """Test that update_list_item sends updated fields to the Graph API."""
    mock_client = _mock_client()
    list_client = lists_mod.GraphList(client=mock_client)

    result = list_client.update_list_item("42", {"Status": "Closed"})

    mock_client.patch.assert_called_once()
    assert result["Title"] == "Updated Item"


def test_get_list_columns_returns_value(env: None) -> None:
    """Test that get_list_columns returns the value array from the columns endpoint."""
    columns = [
        {"name": "Title", "displayName": "Title"},
        {"name": "field_1", "displayName": "Customer Name"},
    ]
    mock_client = _mock_client(return_value={"value": columns})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_columns()

    assert result == columns
    assert mock_client.get.call_count == 2


def test_get_list_columns_can_filter_by_names(env: None) -> None:
    """Test that get_list_columns can request only selected column metadata."""
    columns = [
        {"name": "Title", "displayName": "Título"},
        {"name": "field_1", "displayName": "Customer Name"},
    ]
    mock_client = _mock_client(return_value={"value": columns})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_columns(names=["Title", "field_1"])

    assert result == columns
    call_path = mock_client.get.call_args[0][0]
    assert "columns?$select=name,displayName" in call_path
    assert "name eq 'Title'" in call_path
    assert "name eq 'field_1'" in call_path


def test_get_list_views_returns_value(env: None) -> None:
    """Test that get_list_views returns the value array from the views endpoint."""
    views = [
        {"id": "view-1", "name": "All Items"},
        {"id": "view-2", "name": "Active Only"},
    ]
    mock_client = _mock_client(return_value={"value": views})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_views()

    assert result == views
    assert mock_client.get.call_count == 2


def test_get_list_views_fallback_returns_plain_array(env: None) -> None:
    """Test that get_list_views fallback handles views returned as a plain array.

    Graph API commonly returns expanded collections as plain arrays (not
    wrapped in {"value": [...]}) when using $expand=views. This was the root
    cause of views not being returned when the /views endpoint was unavailable.
    """
    views = [
        {"id": "view-1", "name": "All Items"},
        {"id": "view-2", "name": "Active Only"},
    ]
    mock_client = MagicMock()
    mock_client.get.side_effect = [
        # _get_list_summary
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        # /views endpoint raises HTTPError → triggers fallback
        requests.HTTPError("403 Forbidden"),
        # $expand=views returns views as a plain array (OData expanded collection)
        {"id": "list-abc", "name": "MonitorRNI", "views": views},
    ]

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_views()

    assert result == views


def test_get_list_views_fallback_returns_wrapped_value(env: None) -> None:
    """Test that get_list_views fallback also handles views wrapped in {"value": [...]}."""
    views = [
        {"id": "view-1", "name": "All Items"},
        {"id": "view-2", "name": "Active Only"},
    ]
    mock_client = MagicMock()
    mock_client.get.side_effect = [
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        requests.HTTPError("403 Forbidden"),
        {"id": "list-abc", "name": "MonitorRNI", "views": {"value": views}},
    ]

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_views()

    assert result == views


def test_get_list_views_returns_empty_when_both_endpoints_fail(env: None) -> None:
    """Test that get_list_views returns [] when both the /views and $expand endpoints fail."""
    mock_client = MagicMock()
    mock_client.get.side_effect = [
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        requests.HTTPError("403 Forbidden"),
        requests.HTTPError("403 Forbidden"),
    ]

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_views()

    assert result == []


def test_get_list_view_columns_returns_value(env: None) -> None:
    """Test that get_list_view_columns returns the value array for a specific view."""
    columns = [
        {"name": "Title", "displayName": "Title"},
        {"name": "field_2", "displayName": "Status"},
    ]
    mock_client = _mock_client(return_value={"value": columns})

    list_client = lists_mod.GraphList(client=mock_client)
    result = list_client.get_list_view_columns("view-1")

    assert result == columns
    call_path = mock_client.get.call_args[0][0]
    assert "views/view-1/columns" in call_path
