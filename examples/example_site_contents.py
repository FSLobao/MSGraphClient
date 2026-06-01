"""
example_site_contents.py — Show SharePoint site metadata, drives, lists, and auth context.

Usage:
    uv run examples/example_site_contents.py

Authentication mode is controlled by .env (GRAPH_AUTH_MODE). This script works for:
    - client_credentials
    - delegated
"""

from msgraphclient.auth import GraphAuthorizationError, GraphClient


def _safe(value: object) -> str:
    """Return a printable string fallback for missing values."""
    if value is None:
        return "-"
    text = str(value).strip()
    return text if text else "-"


def _print_auth_details(client: GraphClient) -> None:
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


def main() -> None:
    """Display auth details and the configured site's summary plus available drives and lists."""
    client = GraphClient()
    print(
        "Fetching configured SharePoint site contents and authentication details...\n"
    )

    _print_auth_details(client)

    contents = client.get_site_contents()
    site = contents["site"]
    drives = contents["drives"]
    lists_ = contents["lists"]

    print("Site")
    print(f"  id:          {site.get('id', '-')}")
    print(f"  name:        {site.get('name', '-')}")
    print(f"  displayName: {site.get('displayName', '-')}")
    print(f"  webUrl:      {site.get('webUrl', '-')}")
    print()

    print(f"Drives ({len(drives)})")
    if not drives:
        print("  (none found)")
    for drive in drives:
        print(
            "  - "
            f"{drive.get('name', '(no name)')} | "
            f"id={drive.get('id', '?')} | "
            f"type={drive.get('driveType', '?')}"
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


if __name__ == "__main__":
    try:
        main()
    except GraphAuthorizationError as error:
        print("Authorization failed for current authentication flow.")
        print(GraphClient.format_http_error(error))
