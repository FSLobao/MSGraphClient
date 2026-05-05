"""
example_drive_download.py — Download a file from SharePoint to a local folder.

Set ITEM_ID below to the drive item ID of the file you want to download.
The file will be saved to the downloads/ folder next to this script.

Usage:
    uv run examples/example_drive_download.py
"""

from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

from msgraphtest.drive import list_drive_items, download_file

# ── Configuration ───────────────────────────────────────────────────────────
# Replace with a real drive item ID, or leave empty to use the first file found
ITEM_ID: str = ""
LOCAL_FOLDER: Path = Path(__file__).parent.parent / "downloads"
# ────────────────────────────────────────────────────────────────────────────


def main() -> None:
    """Download a file from the SharePoint drive to the local filesystem.

    If ITEM_ID is not set, downloads the first file found in the drive root.
    Saved files are placed in the downloads/ folder.
    """
    item_id = ITEM_ID

    if not item_id:
        print("No ITEM_ID set — picking the first file from the drive root...")
        items = list_drive_items()
        files = [i for i in items if "folder" not in i]
        if not files:
            print("No files found in drive root.")
            return
        item_id = files[0]["id"]
        filename = files[0]["name"]
        print(f"  Using: {filename} (id={item_id})")
    else:
        filename = f"downloaded_{item_id}"

    dest = LOCAL_FOLDER / filename
    result_path = download_file(item_id, dest)
    print(f"\nFile saved to: {result_path}")


if __name__ == "__main__":
    main()
