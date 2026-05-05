"""
example_drive_read_write.py — Read and then update the text content of a drive item.

Set ITEM_ID to a text-based file in your drive.

Usage:
    uv run examples/example_drive_read_write.py
"""

from dotenv import load_dotenv

load_dotenv()

from msgraphtest.drive import read_file_content, write_file_content

# ── Configuration ───────────────────────────────────────────────────────────
# Replace with a real drive item ID for a text file
ITEM_ID: str = "YOUR_ITEM_ID_HERE"
# ────────────────────────────────────────────────────────────────────────────


def main() -> None:
    """Read, modify, and write back the content of a SharePoint drive file.

    Reads the original content, appends a marker line, and writes the updated
    content back to the same drive item.
    """
    if ITEM_ID == "YOUR_ITEM_ID_HERE":
        print("Please set ITEM_ID in this script to a real drive item ID.")
        return

    print(f"Reading content of item: {ITEM_ID}")
    original = read_file_content(ITEM_ID)
    print("\n--- Original content ---")
    print(original)

    new_content = original + "\n[Appended by msgraphtest example]\n"
    print("\nWriting updated content...")
    result = write_file_content(ITEM_ID, new_content)
    print(f"Update successful. Item ID: {result.get('id')}")

    print("\nVerifying update — reading content again...")
    updated = read_file_content(ITEM_ID)
    print("--- Updated content ---")
    print(updated)


if __name__ == "__main__":
    main()
