"""
example_drive_download.py — Download a file from SharePoint to a local folder.

Set ITEM_ID below to the drive item ID of the file you want to download.
The file will be saved to the examples/downloads/ folder.

Usage:
    uv run examples/example_drive_download.py
"""

from pathlib import Path
from typing import Any
import os


from msgraphclient.auth import GraphClient
from msgraphclient.drive import GraphDrive

# ── Configuration ───────────────────────────────────────────────────────────
# Replace with a real drive item ID, or leave empty to use the first file found
ITEM_ID: str = ""
LOCAL_FOLDER: Path = Path(__file__).parent / "downloads"
# ────────────────────────────────────────────────────────────────────────────


def run_example_drive_download(
    client: GraphClient | None = None,
    drive: GraphDrive | None = None,
    drive_id: str | None = None,
    item_id: str | None = None,
    local_folder: str | Path | None = None,
    show_output: bool = True,
) -> dict[str, Any]:
    """Download a drive file and return context plus local path."""
    resolved_client = client or GraphClient()
    resolved_drive = drive
    if resolved_drive is None:
        resolved_drive_id = drive_id or os.environ["SHAREPOINT_DRIVE_ID"]
        resolved_drive = GraphDrive(drive_id=resolved_drive_id, client=resolved_client)

    target_item_id = (item_id or ITEM_ID).strip()
    target_folder = Path(local_folder) if local_folder else LOCAL_FOLDER

    if not target_item_id:
        if show_output:
            print("No ITEM_ID set - picking the first file from the drive root...")
        items = resolved_drive.ls()
        files = [i for i in items if "folder" not in i]
        if not files:
            if show_output:
                print("No files found in drive root.")
            return {
                "client": resolved_client,
                "authenticator": resolved_client.authenticator,
                "drive": resolved_drive,
                "item_id": "",
                "saved_path": None,
                "success": False,
            }
        target_item_id = str(files[0]["id"])
        filename = str(files[0]["name"])
        if show_output:
            print(f"  Using: {filename} (id={target_item_id})")
    else:
        filename = f"downloaded_{target_item_id}"

    dest = target_folder / filename
    result_path = resolved_drive.download(target_item_id, dest)
    if show_output:
        print(f"\nFile saved to: {result_path}")

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "drive": resolved_drive,
        "item_id": target_item_id,
        "saved_path": result_path,
        "success": True,
    }


def main() -> None:
    """Download a file from the SharePoint drive to the local filesystem."""
    run_example_drive_download(show_output=True)


if __name__ == "__main__":
    main()
