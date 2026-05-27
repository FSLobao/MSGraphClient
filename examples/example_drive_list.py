"""
example_drive_list.py — List the contents of the SharePoint document library root.

Usage:
    uv run examples/example_drive_list.py
"""

from dotenv import load_dotenv
from python.auth import GraphClient
from python.drive import GraphDrive

load_dotenv()


def main() -> None:
    """List and display all items in the root of the configured SharePoint drive.

    Shows the name, type (file or folder), and size of each item.
    """
    client = GraphClient()
    drive = GraphDrive(client=client)

    print("Listing items in the root of the configured drive...\n")
    items = drive.list_drive_items()
    if not items:
        print("(no items found)")
        return
    for item in items:
        kind = "folder" if "folder" in item else "file "
        size = item.get("size", "-")
        print(f"  [{kind}]  {item['name']:<40}  size={size}")
    print(f"\nTotal: {len(items)} item(s)")


if __name__ == "__main__":
    main()
