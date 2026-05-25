"""
site.py — SharePoint site discovery helpers via Microsoft Graph.

Provides helpers to fetch metadata and available resources for the
configured SharePoint site identified by SHAREPOINT_SITE_ID.
"""

from __future__ import annotations

import os

from dotenv import load_dotenv

from msgraphtest.graph_client import GraphClient

load_dotenv()


def _site_id() -> str:
    """Retrieve the SharePoint site ID from environment configuration.

    Reads SHAREPOINT_SITE_ID from environment variables (typically set
    via .env file).

    Returns:
        The SharePoint site ID string.

    Raises:
        EnvironmentError: If SHAREPOINT_SITE_ID is not set or is empty.
    """
    site_id = os.environ.get("SHAREPOINT_SITE_ID", "")
    if not site_id:
        raise EnvironmentError("SHAREPOINT_SITE_ID environment variable is not set.")
    return site_id


def get_site_summary() -> dict:
    """Return metadata for the configured SharePoint site.

    Returns:
        A Graph site dict with selected metadata fields.
    """
    client = GraphClient()
    site_id = _site_id()
    select = (
        "id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime"
    )
    return client.get(f"/sites/{site_id}?$select={select}")


def list_site_drives() -> list[dict]:
    """Return all document libraries (drives) for the configured site."""
    client = GraphClient()
    site_id = _site_id()
    data = client.get(f"/sites/{site_id}/drives?$select=id,name,webUrl,driveType")
    return data.get("value", [])


def list_site_lists() -> list[dict]:
    """Return all SharePoint lists for the configured site."""
    client = GraphClient()
    site_id = _site_id()
    data = client.get(f"/sites/{site_id}/lists?$select=id,name,displayName,webUrl")
    return data.get("value", [])


def get_site_contents() -> dict:
    """Return a consolidated snapshot of site metadata, drives, and lists.

    Returns:
        A dict containing:
            - ``site``: site metadata dict
            - ``drives``: list of drive dicts
            - ``lists``: list of list dicts
    """
    return {
        "site": get_site_summary(),
        "drives": list_site_drives(),
        "lists": list_site_lists(),
    }
