"""
drive.py — SharePoint document library (drive) operations via Microsoft Graph.

All functions operate against a specific drive identified by SHAREPOINT_DRIVE_ID.
The site is identified by SHAREPOINT_SITE_ID.

Covered operations:
    list_drive_items    — list the children of a folder (default: root)
    download_file       — download a drive item to a local path
    upload_file         — upload a local file to the drive
    read_file_content   — return the text content of a drive item
    write_file_content  — overwrite a drive item with new text content
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

from msgraphtest.graph_client import GraphClient

load_dotenv()


def _drive_id() -> str:
    """Retrieve the SharePoint drive ID from environment configuration.

    Reads SHAREPOINT_DRIVE_ID from environment variables (typically set
    via .env file).

    Returns:
        The SharePoint drive ID string.

    Raises:
        EnvironmentError: If SHAREPOINT_DRIVE_ID is not set or is empty.
    """
    drive_id = os.environ.get("SHAREPOINT_DRIVE_ID", "")
    if not drive_id:
        raise EnvironmentError("SHAREPOINT_DRIVE_ID environment variable is not set.")
    return drive_id


def list_drive_items(folder_path: str = "root") -> list[dict]:
    """Return the children of *folder_path* in the configured drive.

    Args:
        folder_path: A drive path string such as ``"root"`` or
            ``"root:/Documents/Reports:"``.  Defaults to ``"root"``.

    Returns:
        A list of Graph driveItem dicts.
    """
    client = GraphClient()
    drive_id = _drive_id()
    path = f"/drives/{drive_id}/items/{folder_path}/children"
    data = client.get(path)
    return data.get("value", [])


def download_file(item_id: str, local_path: str | Path) -> Path:
    """Download a drive item to *local_path*.

    Args:
        item_id:    The drive item ID.
        local_path: Destination file path on the local filesystem.

    Returns:
        The resolved local :class:`~pathlib.Path`.
    """
    client = GraphClient()
    drive_id = _drive_id()
    raw = client.get_raw(f"/drives/{drive_id}/items/{item_id}/content")
    dest = Path(local_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(raw)
    return dest.resolve()


def upload_file(
    local_path: str | Path, remote_folder: str = "root", remote_name: str | None = None
) -> dict:
    """Upload a local file to *remote_folder* in the configured drive.

    Uses the simple upload endpoint (files up to 4 MB).

    Args:
        local_path:    Path to the local file to upload.
        remote_folder: Target folder expressed as a drive item path,
            e.g. ``"root:/Documents:"``.  Defaults to drive root.
        remote_name:   Desired filename in the drive.  Defaults to the
            local filename.

    Returns:
        The Graph driveItem dict for the uploaded file.
    """
    client = GraphClient()
    drive_id = _drive_id()
    src = Path(local_path)
    name = remote_name or src.name
    data = src.read_bytes()
    path = f"/drives/{drive_id}/items/{remote_folder}:/{name}:/content"
    return client.put_bytes(path, data)


def read_file_content(item_id: str, encoding: str = "utf-8") -> str:
    """Return the decoded text content of a drive item.

    Args:
        item_id:  The drive item ID.
        encoding: Text encoding to use when decoding the bytes.

    Returns:
        The file content as a string.
    """
    client = GraphClient()
    drive_id = _drive_id()
    raw = client.get_raw(f"/drives/{drive_id}/items/{item_id}/content")
    return raw.decode(encoding)


def write_file_content(item_id: str, content: str, encoding: str = "utf-8") -> dict:
    """Overwrite an existing drive item with new text *content*.

    Args:
        item_id:  The drive item ID to overwrite.
        content:  New text content.
        encoding: Encoding used to convert the string to bytes.

    Returns:
        The updated Graph driveItem dict.
    """
    client = GraphClient()
    drive_id = _drive_id()
    data = content.encode(encoding)
    return client.put_bytes(
        f"/drives/{drive_id}/items/{item_id}/content",
        data,
        content_type="text/plain",
    )
