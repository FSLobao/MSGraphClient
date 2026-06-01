"""
example_list_create.py — Create a new item in a SharePoint list.

Edit ITEM_FIELDS to set the column values for the new item.

Usage:
    uv run examples/example_list_create.py
"""

import os
from typing import Any

from msgraphclient.auth import GraphClient
from msgraphclient.lists import GraphList

# ── Configuration ───────────────────────────────────────────────────────────
# Adjust these fields to match your list's columns
ITEM_FIELDS: dict = {
    "Title": "Test item created by python",
}
# ────────────────────────────────────────────────────────────────────────────


def run_example_list_create(
    client: GraphClient | None = None,
    list_client: GraphList | None = None,
    list_id: str | None = None,
    item_fields: dict[str, Any] | None = None,
    show_output: bool = True,
) -> dict[str, Any]:
    """Create a list item and return chainable context with the result."""
    resolved_client = client or GraphClient()
    resolved_list_client = list_client
    if resolved_list_client is None:
        resolved_list_id = list_id or os.environ["SHAREPOINT_LIST_ID"]
        resolved_list_client = GraphList(
            list_id=resolved_list_id, client=resolved_client
        )

    fields = dict(item_fields or ITEM_FIELDS)
    if show_output:
        print(f"Creating new list item with fields: {fields}")
    result = resolved_list_client.save_item(fields)

    if show_output:
        print("\nItem created successfully!")
        print(f"  ID     : {result.get('_id')}")
        print(f"  Fields : {result}")

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "list_client": resolved_list_client,
        "created_item": result,
        "item_fields": fields,
        "success": True,
    }


def main() -> None:
    """Create a new item in the configured SharePoint list."""
    run_example_list_create(show_output=True)


if __name__ == "__main__":
    main()
