"""Tests for drive.py"""

import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path

import msgraphtest.drive as drive_mod


@pytest.fixture()
def env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Set up environment variables required for SharePoint drive operations."""
    monkeypatch.setenv("SHAREPOINT_DRIVE_ID", "drive-abc")
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")


def _mock_client(return_value: dict | None = None, raw_bytes: bytes = b"") -> MagicMock:
    """Create a mock GraphClient instance for testing.

    Args:
        return_value: Return value for get() calls. Defaults to empty dict.
        raw_bytes: Raw bytes to return for get_raw() calls.

    Returns:
        A MagicMock object configured to simulate GraphClient behavior.
    """
    client = MagicMock()
    client.get.return_value = return_value or {}
    client.get_raw.return_value = raw_bytes
    client.put_bytes.return_value = {"id": "item-1", "name": "file.txt"}
    return client


def test_list_drive_items_returns_value(env: None) -> None:
    """Test that list_drive_items returns the value array from API response."""
    items = [{"name": "file1.txt"}, {"name": "file2.txt"}]
    mock_client = _mock_client(return_value={"value": items})

    with patch.object(drive_mod, "GraphClient", return_value=mock_client):
        result = drive_mod.list_drive_items()

    assert result == items
    mock_client.get.assert_called_once()


def test_list_drive_items_missing_drive_id(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that list_drive_items raises EnvironmentError when SHAREPOINT_DRIVE_ID is missing."""
    monkeypatch.delenv("SHAREPOINT_DRIVE_ID", raising=False)

    with patch.object(drive_mod, "GraphClient"):
        with pytest.raises(EnvironmentError, match="SHAREPOINT_DRIVE_ID"):
            drive_mod.list_drive_items()


def test_download_file(env: None, tmp_path: Path) -> None:
    """Test that download_file correctly fetches and writes a file to local disk."""
    mock_client = _mock_client(raw_bytes=b"file content")

    with patch.object(drive_mod, "GraphClient", return_value=mock_client):
        dest = tmp_path / "downloaded.txt"
        result = drive_mod.download_file("item-123", dest)

    assert result == dest.resolve()
    assert dest.read_bytes() == b"file content"


def test_upload_file(env: None, tmp_path: Path) -> None:
    """Test that upload_file sends file bytes to the Graph API."""
    src = tmp_path / "upload_me.txt"
    src.write_bytes(b"hello world")
    mock_client = _mock_client()

    with patch.object(drive_mod, "GraphClient", return_value=mock_client):
        result = drive_mod.upload_file(src)

    mock_client.put_bytes.assert_called_once()
    assert result["name"] == "file.txt"


def test_read_file_content(env: None) -> None:
    """Test that read_file_content decodes binary response as UTF-8 text."""
    mock_client = _mock_client(raw_bytes="Hello, Graph!".encode("utf-8"))

    with patch.object(drive_mod, "GraphClient", return_value=mock_client):
        content = drive_mod.read_file_content("item-456")

    assert content == "Hello, Graph!"


def test_write_file_content(env: None) -> None:
    """Test that write_file_content encodes text and sends to Graph API."""
    mock_client = _mock_client()
    mock_client.put_bytes.return_value = {"id": "item-456"}

    with patch.object(drive_mod, "GraphClient", return_value=mock_client):
        result = drive_mod.write_file_content("item-456", "updated content")

    mock_client.put_bytes.assert_called_once()
    assert result["id"] == "item-456"
