"""Microsoft Graph HTTP client and authorization error.

This module is intentionally kept separate from authentication logic
so that HTTP-level concerns (request dispatch, error formatting) are
decoupled from token acquisition.

The ``GraphAuthenticator`` class lives in :mod:`python.auth` and is
imported lazily inside :meth:`GraphClient.__init__` to avoid a circular
dependency (``GraphAuthenticator`` creates ``GraphClient`` instances at
runtime and vice-versa).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

import requests

if TYPE_CHECKING:
    from python.auth import GraphAuthenticator

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

__all__ = ["GraphAuthorizationError", "GraphClient"]


class GraphAuthorizationError(requests.HTTPError):
    """HTTP error raised when caller lacks authorization to a Graph resource."""


class GraphClient:
    """Minimal Microsoft Graph API client.

    Public methods:
        - format_http_error
        - get
        - post
        - patch
        - put_bytes
        - get_raw

    Public attributes:
        - authenticator

    Internal methods (implementation detail):
        - _extract_graph_error_detail
        - _raise_for_status
    """

    @staticmethod
    def format_http_error(error: requests.HTTPError) -> str:
        """Return a clean, user-facing message for an HTTP error."""
        response = error.response
        if response is None:
            return f"HTTP error: {error}"

        method = response.request.method if response.request else "HTTP"
        url = response.url or "<unknown-url>"
        status = response.status_code
        reason = response.reason or ""
        base = f"{method} {url} failed with {status} {reason}".strip()

        detail = GraphClient._extract_graph_error_detail(response)
        if detail:
            return f"{base}. Detail: {detail}"
        return base

    @staticmethod
    def _extract_graph_error_detail(response: requests.Response) -> str | None:
        """Extract Graph error code/message from an HTTP response, if available."""
        try:
            payload = response.json()
        except ValueError:
            text = response.text.strip()
            return text or None

        if not isinstance(payload, dict):
            return None

        error_obj = payload.get("error")
        if isinstance(error_obj, dict):
            code = str(error_obj.get("code", "")).strip()
            message = str(error_obj.get("message", "")).strip()
            if code and message:
                return f"{code}: {message}"
            if message:
                return message
            if code:
                return code

        return None

    @staticmethod
    def _raise_for_status(response: requests.Response) -> None:
        """Raise enriched HTTP exceptions, specializing authorization failures."""
        try:
            response.raise_for_status()
            return
        except requests.HTTPError as error:
            message = GraphClient.format_http_error(error)
            status = response.status_code
            if status in (401, 403):
                raise GraphAuthorizationError(
                    f"Authorization error: {message}",
                    response=response,
                    request=response.request,
                ) from error

            raise requests.HTTPError(
                message,
                response=response,
                request=response.request,
            ) from error

    def __init__(
        self,
        token: str | None = None,
        authenticator: GraphAuthenticator | None = None,
        sharepoint_site_id: str | None = None,
        auth_mode: str | None = None,
    ) -> None:
        """Initialize Graph client and ensure an attached GraphAuthenticator.

        Args:
            token: Optional explicit bearer token.
            authenticator: Optional pre-built GraphAuthenticator to reuse.
            sharepoint_site_id: Optional site id forwarded to authenticator init.
            auth_mode: Optional auth mode override (client_credentials | delegated).
        """
        # Lazy import breaks the circular dependency with python.auth.
        from python.auth import GraphAuthenticator as _GraphAuthenticator

        if authenticator is None:
            self.authenticator = _GraphAuthenticator(
                sharepoint_site_id=sharepoint_site_id,
                token=token,
                create_client=False,
                auth_mode=auth_mode,
            )
        else:
            self.authenticator = authenticator
            if sharepoint_site_id:
                self.authenticator.sharepoint_site_id = sharepoint_site_id

        self._token: str = (
            token
            or self.authenticator.token
            or _GraphAuthenticator._acquire_token_from_env_internal(
                auth_mode=self.authenticator.auth_mode
            )
        )
        self.authenticator.token = self._token

        self._session = requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {self._token}",
                "Accept": "application/json",
            }
        )

        # Bind authenticator to this concrete Graph client instance.
        self.authenticator.client = self
        self.authenticator._client = self
        if self.authenticator.sharepoint_site_id and not self.authenticator.site_data:
            self.authenticator._load_site_summary()

    def get(self, path: str, **kwargs: Any) -> dict:
        """Make a GET request to the Graph API."""
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.get(url, **kwargs)
        self._raise_for_status(response)
        return response.json()

    def post(self, path: str, json: dict, **kwargs: Any) -> dict:
        """Make a POST request to the Graph API."""
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.post(url, json=json, **kwargs)
        self._raise_for_status(response)
        return response.json()

    def patch(self, path: str, json: dict, **kwargs: Any) -> dict:
        """Make a PATCH request to the Graph API."""
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.patch(url, json=json, **kwargs)
        self._raise_for_status(response)
        return response.json()

    def put_bytes(
        self,
        path: str,
        data: bytes,
        content_type: str = "application/octet-stream",
        **kwargs: Any,
    ) -> dict:
        """Make a PUT request to the Graph API with binary data."""
        url = f"{GRAPH_BASE_URL}{path}"
        headers = {"Content-Type": content_type}
        response = self._session.put(url, data=data, headers=headers, **kwargs)
        self._raise_for_status(response)
        return response.json()

    def get_raw(self, path: str, **kwargs: Any) -> bytes:
        """Make a GET request and return raw binary response body."""
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.get(url, **kwargs)
        self._raise_for_status(response)
        return response.content
