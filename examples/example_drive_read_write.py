"""
example_drive_read_write.py — Read and then update the text content of a drive item.

Set DRIVE_ITEM_ID in .env to a text-based file in your drive.

Usage:
    uv run examples/example_drive_read_write.py
"""

import os
from typing import Any


from msgraphclient.auth import GraphClient
from msgraphclient.drive import GraphDrive

# ── Configuration ───────────────────────────────────────────────────────────
# Set DRIVE_ITEM_ID in .env with a real drive item ID for a text file
ITEM_ID: str = os.getenv("DRIVE_ITEM_ID", "").strip()
# ────────────────────────────────────────────────────────────────────────────


def run_example_drive_read_write(
    client: GraphClient | None = None,
    drive: GraphDrive | None = None,
    drive_id: str | None = None,
    item_id: str | None = None,
    append_suffix: str = "\n[Appended by python example]\n",
    show_output: bool = True,
) -> dict[str, Any]:
    """Read, update, and re-read a text file, returning chainable context."""
    resolved_client = client or GraphClient()
    resolved_drive = drive
    if resolved_drive is None:
        resolved_drive_id = drive_id or os.environ["SHAREPOINT_DRIVE_ID"]
        resolved_drive = GraphDrive(drive_id=resolved_drive_id, client=resolved_client)

    target_item_id = (item_id or ITEM_ID).strip()
    if not target_item_id:
        if show_output:
            print("Please set DRIVE_ITEM_ID in .env or pass item_id to the function.")
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "drive": resolved_drive,
            "item_id": "",
            "original_content": None,
            "updated_content": None,
            "write_result": None,
            "success": False,
        }

    if show_output:
        print(f"Reading content of item: {target_item_id}")
    original = resolved_drive.read_file_content(target_item_id)
    if show_output:
        print("\n--- Original content ---")
        print(original)

    new_content = original + append_suffix
    if show_output:
        print("\nWriting updated content...")
    result = resolved_drive.write_file_content(target_item_id, new_content)
    if show_output:
        print(f"Update successful. Item ID: {result.get('id')}")

    if show_output:
        print("\nVerifying update - reading content again...")
    updated = resolved_drive.read_file_content(target_item_id)
    if show_output:
        print("--- Updated content ---")
        print(updated)

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "drive": resolved_drive,
        "item_id": target_item_id,
        "original_content": original,
        "updated_content": updated,
        "write_result": result,
        "success": True,
    }


def main() -> None:
    """Read, modify, and write back the content of a SharePoint drive file."""
    run_example_drive_read_write(show_output=True)


if __name__ == "__main__":
    main()
