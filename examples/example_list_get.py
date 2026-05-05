"""
example_list_get.py — Retrieve and display all items from a SharePoint list.

Usage:
    uv run examples/example_list_get.py
"""

from dotenv import load_dotenv

load_dotenv()

from msgraphtest.lists import get_list_items


def main() -> None:
    """Retrieve and display all items from the configured SharePoint list.

    Shows the ID and Title field of each list item.
    """
    print("Fetching items from the configured SharePoint list...\n")
    items = get_list_items()
    if not items:
        print("(no items found)")
        return
    for item in items:
        fields = item.get("fields", {})
        title = fields.get("Title", "(no title)")
        item_id = item.get("id", "?")
        print(f"  [{item_id:>4}]  {title}")
    print(f"\nTotal: {len(items)} item(s)")


if __name__ == "__main__":
    main()
