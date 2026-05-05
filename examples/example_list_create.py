"""
example_list_create.py — Create a new item in a SharePoint list.

Edit ITEM_FIELDS to set the column values for the new item.

Usage:
    uv run examples/example_list_create.py
"""

from dotenv import load_dotenv

load_dotenv()

from msgraphtest.lists import create_list_item

# ── Configuration ───────────────────────────────────────────────────────────
# Adjust these fields to match your list's columns
ITEM_FIELDS: dict = {
    "Title": "Test item created by msgraphtest",
}
# ────────────────────────────────────────────────────────────────────────────


def main() -> None:
    """Create a new item in the configured SharePoint list.

    Uses the fields defined in ITEM_FIELDS and displays the created item's
    ID and assigned field values.
    """
    print(f"Creating new list item with fields: {ITEM_FIELDS}")
    result = create_list_item(ITEM_FIELDS)
    print(f"\nItem created successfully!")
    print(f"  ID     : {result.get('id')}")
    print(f"  Fields : {result.get('fields')}")


if __name__ == "__main__":
    main()
