"""
auth.py — MSAL client-credentials authentication helper.

Acquires an access token for the Microsoft Graph API using the
OAuth 2.0 client credentials flow (app-only).  Configuration is
read from environment variables (or a .env file loaded beforehand).

Required environment variables:
    AZURE_TENANT_ID     – Azure AD tenant ID
    AZURE_CLIENT_ID     – App registration client ID
    AZURE_CLIENT_SECRET – App registration client secret

Usage::

    from msgraphtest.auth import get_access_token
    token = get_access_token()
"""

import os

import msal
from dotenv import load_dotenv

load_dotenv()

GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]


def _get_config() -> tuple[str, str, str]:
    """Retrieve and validate Azure AD authentication configuration from environment.

    Reads AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET from
    environment variables (typically set via .env file).

    Returns:
        A tuple of (tenant_id, client_id, client_secret).

    Raises:
        EnvironmentError: If any required environment variable is missing or empty.
    """
    tenant_id = os.environ.get("AZURE_TENANT_ID", "")
    client_id = os.environ.get("AZURE_CLIENT_ID", "")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET", "")
    if not all([tenant_id, client_id, client_secret]):
        raise EnvironmentError(
            "Missing one or more required environment variables: "
            "AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET"
        )
    return tenant_id, client_id, client_secret


def get_access_token() -> str:
    """Return a valid Graph API bearer token using client credentials."""
    tenant_id, client_id, client_secret = _get_config()
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )

    result = app.acquire_token_for_client(scopes=GRAPH_SCOPES)

    if "access_token" not in result:
        error = result.get("error", "unknown")
        description = result.get("error_description", "")
        raise RuntimeError(f"Failed to acquire token: {error} — {description}")

    return result["access_token"]
