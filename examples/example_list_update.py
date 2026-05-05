"""
example_list_update.py — Update fields on an existing SharePoint list item.

Set ITEM_ID to the numeric string ID of the item you want to update,
then edit UPDATE_FIELDS with the columns and new values.

Usage:
    uv run examples/example_list_update.py
"""

from dotenv import load_dotenv

load_dotenv()

from msgraphtest.lists import get_list_items, update_list_item

# ── Configuration ───────────────────────────────────────────────────────────
# Set to the ID of the item to update, or leave empty to update the first item
ITEM_ID: str = ""
UPDATE_FIELDS: dict = {
    "Title": "Updated by msgraphtest",
}
# ────────────────────────────────────────────────────────────────────────────


def main() -> None:
    """Update fields on an existing SharePoint list item.

    If ITEM_ID is not set, updates the first item in the list. Updates the
    fields specified in UPDATE_FIELDS.
    """
    item_id = ITEM_ID

    if not item_id:
        print("No ITEM_ID set — fetching the first list item...")
        items = get_list_items()
        if not items:
            print("No items found in the list.")
            return
        item_id = items[0]["id"]
        print(f"  Using item ID: {item_id}")

    print(f"\nUpdating item {item_id} with: {UPDATE_FIELDS}")
    result = update_list_item(item_id, UPDATE_FIELDS)
    print("\nUpdate successful!")
    print(f"  Updated fields: {result}")


if __name__ == "__main__":
    main()
