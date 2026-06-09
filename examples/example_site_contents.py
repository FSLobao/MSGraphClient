"""
example_site_contents.py — Show SharePoint site metadata, drives, lists, and auth context.

Usage:
    uv run examples/example_site_contents.py

Authentication mode is controlled by .env (GRAPH_AUTH_MODE). This script works for:
    - client_credentials
    - delegated
"""

import os
from typing import Any

from ezspi.auth import AuthorizationError, Client
from ezspi.drive import SPLibrary
from ezspi.lists import SPList


def _safe(value: object) -> str:
    """Return a printable string fallback for missing values."""
    if value is None:
        return "-"
    text = str(value).strip()
    return text if text else "-"


def _print_auth_details(client: Client) -> None:
    """Print non-secret authentication details, including delegated user attributes."""
    auth = client.authenticator

    print("Authentication")
    print(f"  mode:                 {_safe(auth.auth_mode)}")
    print(f"  tenant_id:            {_safe(auth.tenant_id)}")
    print(f"  client_id:            {_safe(auth.client_id)}")
    print(f"  delegated_login_mode: {_safe(auth.delegated_login_mode)}")
    print(
        "  delegated_scopes:     "
        f"{', '.join(auth.delegated_scopes) if auth.delegated_scopes else '-'}"
    )

    if auth.auth_mode == "delegated":
        try:
            me = client.get(
                "/me?$select=id,displayName,userPrincipalName,mail,jobTitle"
            )
            print("  delegated_user:")
            print(f"    id:                {_safe(me.get('id'))}")
            print(f"    displayName:       {_safe(me.get('displayName'))}")
            print(f"    userPrincipalName: {_safe(me.get('userPrincipalName'))}")
            print(f"    mail:              {_safe(me.get('mail'))}")
            print(f"    jobTitle:          {_safe(me.get('jobTitle'))}")
        except Exception as error:  # noqa: BLE001
            print(f"  delegated_user: unavailable ({error})")
    print()


def run_example_site_contents(
    client: Client | None = None,
    drive: SPLibrary | None = None,
    list_client: SPList | None = None,
    drive_id: str | None = None,
    list_id: str | None = None,
    show_output: bool = True,
) -> dict[str, Any]:
    """Return site details and reusable objects for downstream examples."""
    resolved_client = client or Client()

    if show_output:
        print(
            "Fetching configured SharePoint site contents and authentication details...\n"
        )
        _print_auth_details(resolved_client)

    contents = resolved_client.get_site_contents()
    site = contents["site"]
    drives = contents["drives"]
    lists_ = contents["lists"]

    resolved_drive = drive
    if resolved_drive is None:
        resolved_drive_id = drive_id or os.environ.get("SHAREPOINT_DRIVE_ID", "")
        if resolved_drive_id:
            resolved_drive = SPLibrary(
                drive_id=resolved_drive_id, client=resolved_client
            )

    resolved_list_client = list_client
    if resolved_list_client is None:
        resolved_list_id = list_id or os.environ.get("SHAREPOINT_LIST_ID", "")
        if resolved_list_id:
            resolved_list_client = SPList(
                list_id=resolved_list_id,
                client=resolved_client,
            )

    if show_output:
        print("Site")
        print(f"  id:          {site.get('id', '-')}")
        print(f"  name:        {site.get('name', '-')}")
        print(f"  displayName: {site.get('displayName', '-')}")
        print(f"  webUrl:      {site.get('webUrl', '-')}")
        print()

        print(f"Drives ({len(drives)})")
        if not drives:
            print("  (none found)")
        for drive_item in drives:
            print(
                "  - "
                f"{drive_item.get('name', '(no name)')} | "
                f"id={drive_item.get('id', '?')} | "
                f"type={drive_item.get('driveType', '?')}"
            )
        print()

        print(f"Lists ({len(lists_)})")
        if not lists_:
            print("  (none found)")
        for item in lists_:
            print(
                "  - "
                f"{item.get('displayName', item.get('name', '(no name)'))} | "
                f"id={item.get('id', '?')}"
            )

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "site_contents": contents,
        "site": site,
        "drives": drives,
        "lists": lists_,
        "drive": resolved_drive,
        "list_client": resolved_list_client,
    }


def main() -> None:
    """Display auth details and the configured site's summary plus available drives and lists."""
    run_example_site_contents(show_output=True)


if __name__ == "__main__":
    try:
        main()
    except AuthorizationError as error:
        print("Authorization failed for current authentication flow.")
        print(Client.format_http_error(error))

