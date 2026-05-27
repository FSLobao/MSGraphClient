"""example_list_update.py — Perform a typed point update on one list item.

The script mirrors the notebook flow and applies one update payload containing
multiple validated field types when they exist in the target list schema:
text, number, boolean, dateTime, and choice.

Usage:
    uv run examples/example_list_update.py
"""

from datetime import datetime, timezone

from dotenv import load_dotenv

load_dotenv()

from python.auth import GraphClient
from python.lists import GraphList

# Set to the ID of the item to update, or leave empty to update the first item.
ITEM_ID: str = ""


def _build_typed_update(list_client: GraphList) -> dict:
    """Build a point-update payload with one value per supported type."""
    schema_by_name = {
        entry["display_name"]: entry for entry in list_client.get_schema()
    }

    payload: dict = {}
    timestamp = datetime.now(timezone.utc)

    def first_writable(field_type: str) -> str | None:
        for entry in list_client.get_schema():
            if entry.get("read_only"):
                continue
            if entry.get("type") == field_type:
                return str(entry["display_name"])
        return None

    text_col = first_writable("text")
    if text_col:
        payload[text_col] = f"Updated via script at {timestamp.isoformat()}"

    number_col = first_writable("number")
    if number_col:
        payload[number_col] = 123.45

    boolean_col = first_writable("boolean")
    if boolean_col:
        payload[boolean_col] = True

    datetime_col = first_writable("dateTime")
    if datetime_col:
        payload[datetime_col] = timestamp

    choice_col = first_writable("choice")
    if choice_col:
        choices = schema_by_name.get(choice_col, {}).get("choices", [])
        if choices:
            payload[choice_col] = choices[0]

    return payload


def main() -> None:
    """Execute a typed point update on one existing list item."""
    client = GraphClient()
    list_client = GraphList(client=client)

    item_id = ITEM_ID

    if not item_id:
        print("No ITEM_ID set — fetching the first list item...")
        items_df = list_client.get_items_dataframe(include_id=True)
        if items_df.empty:
            print("No items found in the list.")
            return
        item_id = str(items_df.iloc[0]["_id"])
        print(f"  Using item ID: {item_id}")

    typed_update = _build_typed_update(list_client)
    if not typed_update:
        print("No writable fields of supported types were found to update.")
        return

    payload = {"_id": item_id, **typed_update}

    print(f"\nUpdating item {item_id} with typed payload:")
    for key, value in typed_update.items():
        print(f"  - {key}: {value}")

    result = list_client.save_item(payload)

    print("\nUpdate successful!")
    print("  Saved item (display-name format):")
    print(f"  {result}")


if __name__ == "__main__":
    main()
