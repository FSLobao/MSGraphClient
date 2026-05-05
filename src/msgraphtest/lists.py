"""
lists.py — SharePoint list operations via Microsoft Graph.

All functions operate against a specific list identified by
SHAREPOINT_LIST_ID inside the site identified by SHAREPOINT_SITE_ID.

Covered operations:
    get_list_items   — retrieve all items from the list
    create_list_item — create a new list item
    update_list_item — update fields on an existing list item
"""

from __future__ import annotations

import os

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
