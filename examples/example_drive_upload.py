"""Upload a local file to a SharePoint drive, supporting object reuse between examples."""

import os
from pathlib import Path
from typing import Any

from requests import HTTPError

from ezspi.auth import Client
from ezspi.drive import SPLibrary


# ── Configuration ───────────────────────────────────────────────────────────
# Path to the file you want to upload
LOCAL_FILE: Path = Path(__file__).parent / "downloads" / "sample_upload.txt"
# Target folder in the drive, e.g. "root:/Documents:" — defaults to drive root
REMOTE_FOLDER: str = "root"
# ────────────────────────────────────────────────────────────────────────────


def run_example_drive_upload(
    client: Client | None = None,
    drive: SPLibrary | None = None,
    drive_id: str | None = None,
    local_file: str | Path | None = None,
    remote_folder: str = REMOTE_FOLDER,
    remote_name: str | None = None,
    create_sample_if_missing: bool = True,
    show_output: bool = True,
) -> dict[str, Any]:
    """Upload a file and return reusable context plus upload result."""
    resolved_client = client or Client()
    resolved_drive = drive
    if resolved_drive is None:
        resolved_drive_id = drive_id or os.environ["SHAREPOINT_DRIVE_ID"]
        resolved_drive = SPLibrary(drive_id=resolved_drive_id, client=resolved_client)

    source_file = Path(local_file) if local_file else LOCAL_FILE
    if create_sample_if_missing and not source_file.exists():
        source_file.parent.mkdir(parents=True, exist_ok=True)
        source_file.write_text("This is a sample file uploaded by python.\n")
        if show_output:
            print(f"Created sample file: {source_file}")

    if show_output:
        print(f"Uploading {source_file.name} to drive folder '{remote_folder}'...")

    try:
        result = resolved_drive.upload(
            source_file,
            remote_folder=remote_folder,
            remote_name=remote_name,
        )
    except HTTPError as exc:
        if show_output:
            print("\nUpload failed.")
            print(f"  {Client.format_http_error(exc)}")
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "drive": resolved_drive,
            "local_file": source_file,
            "upload_result": None,
            "error": exc,
            "success": False,
        }
    except Exception as exc:  # noqa: BLE001
        if show_output:
            print("\nUpload failed due to an unexpected error.")
            print(f"  {exc}")
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "drive": resolved_drive,
            "local_file": source_file,
            "upload_result": None,
            "error": exc,
            "success": False,
        }

    if show_output:
        print("\nUpload successful!")
        print(f"  Item ID  : {result.get('id')}")
        print(f"  Name     : {result.get('name')}")
        print(f"  Web URL  : {result.get('webUrl')}")

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "drive": resolved_drive,
        "local_file": source_file,
        "upload_result": result,
        "success": True,
    }


def main() -> int:
    """Upload a local file to the SharePoint drive."""
    context = run_example_drive_upload(show_output=True)
    return 0 if context["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

