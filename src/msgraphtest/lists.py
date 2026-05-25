"""
lists.py — SharePoint list operations via Microsoft Graph.

All functions operate against a specific list identified by
SHAREPOINT_LIST_ID inside the site identified by SHAREPOINT_SITE_ID.

Covered operations:
    get_list_views        — list available views (id, name)
    get_list_view_columns — retrieve columns visible in a specific view
    get_list_columns      — retrieve all column definitions (name → displayName mapping)
    get_list_items        — retrieve all items from the list
    create_list_item      — create a new list item
    update_list_item      — update fields on an existing list item
"""

from __future__ import annotations

import os

import requests
from dotenv import load_dotenv

from msgraphtest.graph_client import GraphClient

load_dotenv()


def _site_id() -> str:
    """Retrieve the SharePoint site ID from environment configuration.

    Reads SHAREPOINT_SITE_ID from environment variables (typically set
    via .env file).

    Returns:
        The SharePoint site ID string.

    Raises:
        EnvironmentError: If SHAREPOINT_SITE_ID is not set or is empty.
    """
    site_id = os.environ.get("SHAREPOINT_SITE_ID", "")
    if not site_id:
        raise EnvironmentError("SHAREPOINT_SITE_ID environment variable is not set.")
    return site_id


def _list_id() -> str:
    """Retrieve the SharePoint list ID from environment configuration.

    Reads SHAREPOINT_LIST_ID from environment variables (typically set
    via .env file).

    Returns:
        The SharePoint list ID string.

    Raises:
        EnvironmentError: If SHAREPOINT_LIST_ID is not set or is empty.
    """
    list_id = os.environ.get("SHAREPOINT_LIST_ID", "")
    if not list_id:
        raise EnvironmentError("SHAREPOINT_LIST_ID environment variable is not set.")
    return list_id


def get_list_views() -> list[dict]:
    """Retrieve all views defined for the configured SharePoint list.

    Tries the dedicated ``/views`` endpoint first.  If that returns an HTTP
    error (common with ``Sites.Selected`` on certain list types), falls back to
    fetching views as an expanded property of the list resource via
    ``?$expand=views``.

    Returns:
        A list of view dicts, each containing at minimum ``id`` and ``name``.

    Raises:
        requests.HTTPError: If both API calls fail.
    """
    client = GraphClient()
    site_id = _site_id()
    list_id = _list_id()
    try:
        data = client.get(f"/sites/{site_id}/lists/{list_id}/views")
        return data.get("value", [])
    except requests.HTTPError:
        # Fallback: some list types do not expose the dedicated /views
        # sub-endpoint (returns 400) but do support $expand on the list.
        data = client.get(f"/sites/{site_id}/lists/{list_id}?$expand=views")
        views_block = data.get("views", {})
        if isinstance(views_block, dict):
            return views_block.get("value", [])
        return []


def get_list_view_columns(view_id: str) -> list[dict]:
    """Retrieve the column definitions visible in a specific list view.

    The returned dicts have the same shape as those from :func:`get_list_columns`
    (``name`` and ``displayName``), so the same rename helpers work for both.

    Args:
        view_id: The GUID of the view, as returned by :func:`get_list_views`.

    Returns:
        A list of column definition dicts for the requested view.
    """
    client = GraphClient()
    site_id = _site_id()
    list_id = _list_id()
    data = client.get(f"/sites/{site_id}/lists/{list_id}/views/{view_id}/columns")
    return data.get("value", [])


def get_list_columns() -> list[dict]:
    """Retrieve column definitions for the configured SharePoint list.

    Each dict contains at minimum:
      - ``name``        — internal field name used in item ``fields`` (e.g. ``"field_1"``).
      - ``displayName`` — human-readable label shown in SharePoint.

    Returns:
        A list of column definition dicts.
    """
    client = GraphClient()
    site_id = _site_id()
    list_id = _list_id()
    data = client.get(f"/sites/{site_id}/lists/{list_id}/columns")
    return data.get("value", [])


def get_list_items(select: list[str] | None = None) -> list[dict]:
    """Retrieve all items from the configured SharePoint list.

    Args:
        select: Optional list of field names to include in each item
            (e.g. ``["Title", "Status", "id"]``).  If *None*, all
            fields are returned.

    Returns:
        A list of listItem dicts.  The ``fields`` key of each dict
        contains the column values.
    """
    client = GraphClient()
    site_id = _site_id()
    list_id = _list_id()
    path = f"/sites/{site_id}/lists/{list_id}/items?expand=fields"
    if select:
        path += f"&$select={','.join(select)}"
    data = client.get(path)
    return data.get("value", [])


def create_list_item(fields: dict) -> dict:
    """Create a new item in the configured SharePoint list.

    Args:
        fields: A dict of column name → value pairs for the new item,
            e.g. ``{"Title": "My new item", "Status": "Active"}``.

    Returns:
        The Graph listItem dict for the newly created item (includes
        the assigned ``id``).
    """
    client = GraphClient()
    site_id = _site_id()
    list_id = _list_id()
    payload = {"fields": fields}
    return client.post(f"/sites/{site_id}/lists/{list_id}/items", json=payload)


def update_list_item(item_id: str, fields: dict) -> dict:
    """Update fields on an existing list item.

    Args:
        item_id: The string ID of the list item to update.
        fields:  A dict of column name → new value pairs.

    Returns:
        The updated Graph listItem fields dict.
    """
    client = GraphClient()
    site_id = _site_id()
    list_id = _list_id()
    return client.patch(
        f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields",
        json=fields,
    )
