"""
lists.py — SharePoint list operations via Microsoft Graph.

The primary API is the ``GraphList`` class, which validates configuration and
tests list access on initialization.

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

from msgraphtest.auth import GraphClient

load_dotenv()


class GraphList:
    """SharePoint list operations backed by Microsoft Graph.

    On initialization, resolves the site ID from
    ``client.authenticator.sharepoint_site_id``, then from ``SHAREPOINT_SITE_ID``.
    It also validates ``SHAREPOINT_LIST_ID``, creates/reuses a Graph client,
    and fetches basic list metadata to confirm access.
    """

    def __init__(
        self,
        list_id: str | None = None,
        client: GraphClient | None = None,
    ) -> None:
        """Initialize list operations with optional injected configuration.

        Args:
            list_id: Optional SharePoint list ID. If omitted, reads from .env.
            client: Optional pre-configured GraphClient instance.
        """
        self.client = client or GraphClient()
        resolved_site_id = (
            self._site_id_from_client(self.client) or self._site_id_from_env()
        )
        self.site_id: str = resolved_site_id
        self.list_id: str = list_id or self._list_id_from_env()

        # Public list attributes populated from Graph metadata.
        self.list_info: dict = self._get_list_summary()
        self.list_graph_id: str = str(self.list_info.get("id", ""))
        self.list_name: str = str(self.list_info.get("name", ""))
        self.list_display_name: str = str(self.list_info.get("displayName", ""))
        self.list_web_url: str = str(self.list_info.get("webUrl", ""))

    def _site_id_from_env(self) -> str:
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
            raise EnvironmentError(
                "SHAREPOINT_SITE_ID environment variable is not set."
            )
        return site_id

    @staticmethod
    def _site_id_from_client(client: GraphClient) -> str:
        """Return site id from a client's authenticator when available."""
        authenticator = getattr(client, "authenticator", None)
        site_id = getattr(authenticator, "sharepoint_site_id", None)
        if isinstance(site_id, str):
            return site_id
        return ""

    def _list_id_from_env(self) -> str:
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
            raise EnvironmentError(
                "SHAREPOINT_LIST_ID environment variable is not set."
            )
        return list_id

    def _get_list_summary(self) -> dict:
        """Return basic metadata for the configured list."""
        select = "id,name,displayName,webUrl"
        return self.client.get(
            f"/sites/{self.site_id}/lists/{self.list_id}?$select={select}"
        )

    def get_list_views(self) -> list[dict]:
        """Retrieve all views defined for the configured SharePoint list.

        Tries the dedicated ``/views`` endpoint first. If that returns an HTTP
        error (common with ``Sites.Selected`` on certain list types), falls back to
        fetching views as an expanded property of the list resource via
        ``?$expand=views``. If both fail, returns an empty list (some list types
        do not support views).

        Returns:
            A list of view dicts, each containing at minimum ``id`` and ``name``.
            Returns empty list if views are not available for this list type.
        """
        try:
            data = self.client.get(f"/sites/{self.site_id}/lists/{self.list_id}/views")
            return data.get("value", [])
        except requests.HTTPError:
            try:
                data = self.client.get(
                    f"/sites/{self.site_id}/lists/{self.list_id}?$expand=views"
                )
                views_block = data.get("views", {})
                if isinstance(views_block, dict):
                    return views_block.get("value", [])
                return []
            except requests.HTTPError:
                # Some list types (e.g., tasks) don't support views; return empty
                return []

    def get_list_view_columns(self, view_id: str) -> list[dict]:
        """Retrieve the column definitions visible in a specific list view."""
        data = self.client.get(
            f"/sites/{self.site_id}/lists/{self.list_id}/views/{view_id}/columns"
        )
        return data.get("value", [])

    @staticmethod
    def _escape_odata_string(value: str) -> str:
        """Escape single quotes in OData string literals."""
        return value.replace("'", "''")

    def get_list_columns(self, names: list[str] | None = None) -> list[dict]:
        """Retrieve column definitions for the configured SharePoint list.

        Args:
            names: Optional internal column names to limit metadata retrieval.
                When provided, requests only ``name`` and ``displayName`` for
                the designated columns.
        """
        if names:
            unique_names = list(dict.fromkeys(name for name in names if name))
            if not unique_names:
                return []

            names_filter = " or ".join(
                f"name eq '{self._escape_odata_string(name)}'" for name in unique_names
            )
            path = (
                f"/sites/{self.site_id}/lists/{self.list_id}/columns"
                f"?$select=name,displayName&$filter={names_filter}"
            )
        else:
            path = f"/sites/{self.site_id}/lists/{self.list_id}/columns"

        data = self.client.get(path)
        return data.get("value", [])

    @staticmethod
    def _normalize_selected_fields(
        select: list[str] | None,
        include_title: bool,
    ) -> list[str]:
        """Return an ordered list of unique field names for Graph selection."""
        selected_fields: list[str] = []

        if include_title:
            selected_fields.append("Title")

        for field_name in select or []:
            if field_name and field_name not in selected_fields:
                selected_fields.append(field_name)

        return selected_fields

    def get_list_items(
        self,
        select: list[str] | None = None,
        *,
        include_title: bool = False,
        fields_only: bool = False,
        include_item_id: bool = False,
    ) -> list[dict]:
        """Retrieve items from the configured SharePoint list.

        Args:
            select: Optional internal field names to request from Graph. When
                provided, the request uses ``expand=fields(select=...)`` so the
                response omits unselected SharePoint field metadata.
            include_title: When ``True``, include the ``Title`` field in the
                Graph selection even if it is not present in ``select``. In
                ``fields_only`` mode without ``select``, controls whether
                ``Title`` is returned alongside ``field_*`` keys.
            fields_only: When ``True``, return only each item's ``fields`` block
                instead of the full Graph ``listItem`` envelope.
            include_item_id: When ``True`` together with ``fields_only``, keep
                each item's Graph ``id`` in the returned record.
        """
        has_select = bool(select)
        selected_fields = self._normalize_selected_fields(select, include_title)

        expand = "fields"
        if has_select and selected_fields:
            expand = f"fields(select={','.join(selected_fields)})"

        path = f"/sites/{self.site_id}/lists/{self.list_id}/items?expand={expand}"
        data = self.client.get(path)
        items = data.get("value", [])

        if fields_only:
            normalized_items: list[dict] = []
            for item in items:
                if not isinstance(item, dict):
                    continue
                fields = item.get("fields", {})
                if not isinstance(fields, dict):
                    continue

                if has_select:
                    row = {
                        key: value
                        for key, value in fields.items()
                        if key in selected_fields
                    }
                else:
                    row = {
                        key: value
                        for key, value in fields.items()
                        if isinstance(key, str) and key.startswith("field_")
                    }
                    if include_title and "Title" in fields:
                        row["Title"] = fields["Title"]

                if include_item_id:
                    row["id"] = str(item.get("id", ""))
                normalized_items.append(row)
            return normalized_items
        return items

    def create_list_item(self, fields: dict) -> dict:
        """Create a new item in the configured SharePoint list."""
        payload = {"fields": fields}
        return self.client.post(
            f"/sites/{self.site_id}/lists/{self.list_id}/items", json=payload
        )

    def update_list_item(self, item_id: str, fields: dict) -> dict:
        """Update fields on an existing list item."""
        return self.client.patch(
            f"/sites/{self.site_id}/lists/{self.list_id}/items/{item_id}/fields",
            json=fields,
        )
