"""List items in a SharePoint drive, supporting object reuse between examples."""

import os
from typing import Any

from msgraphclient.auth import GraphClient
from msgraphclient.drive import GraphDrive


def run_example_drive_list(
    client: GraphClient | None = None,
    drive: GraphDrive | None = None,
    drive_id: str | None = None,
    folder_path: str = "/",
    show_output: bool = True,
) -> dict[str, Any]:
    """List drive items and return reusable context and result data."""
    resolved_client = client or GraphClient()
    resolved_drive = drive
    if resolved_drive is None:
        resolved_drive_id = drive_id or os.environ["SHAREPOINT_DRIVE_ID"]
        resolved_drive = GraphDrive(drive_id=resolved_drive_id, client=resolved_client)

    resolved_drive.cd(folder_path)

    if show_output:
        print(f"Listing items in drive folder '{resolved_drive.pwd()}'...\n")
    items = resolved_drive.ls()

    if show_output:
        if not items:
            print("(no items found)")
        for item in items:
            kind = "folder" if "folder" in item else "file "
            size = item.get("size", "-")
            print(f"  [{kind}]  {item['name']:<40}  size={size}")
        print(f"\nTotal: {len(items)} item(s)")

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "drive": resolved_drive,
        "drive_items": items,
        "folder_path": resolved_drive.pwd(),
    }


def main() -> None:
    """List and display all items in the root of the configured SharePoint drive."""
    run_example_drive_list(show_output=True)


if __name__ == "__main__":
    main()
