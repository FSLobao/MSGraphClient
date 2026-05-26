"""Authentication helpers for ArcGIS REST and WMS calls."""

from __future__ import annotations

import os
import time
from dataclasses import dataclass

import requests


class ArcGISAuthError(RuntimeError):
    """Raised when ArcGIS authentication configuration or token fetch fails."""


class ArcGISTokenProvider:
    """Simple interface for providers that return an ArcGIS access token."""

    def get_token(self, session: requests.Session | None = None) -> str:
        raise NotImplementedError


@dataclass
class ApiKeyAuth(ArcGISTokenProvider):
    """Token provider backed by an ArcGIS API key."""

    api_key: str | None = None
    env_var: str = "ARCGIS_API_KEY"

    def get_token(self, session: requests.Session | None = None) -> str:
        token = (self.api_key or os.getenv(self.env_var, "")).strip()
        if not token:
            raise ArcGISAuthError(
                f"Missing ArcGIS API key. Set {self.env_var} in .env or pass api_key."
            )
        return token


@dataclass
class UserTokenAuth(ArcGISTokenProvider):
    """Token provider for user-authenticated OAuth tokens captured elsewhere."""

    token: str | None = None
    env_var: str = "ARCGIS_OAUTH_TOKEN"

    def get_token(self, session: requests.Session | None = None) -> str:
        value = (self.token or os.getenv(self.env_var, "")).strip()
        if not value:
            raise ArcGISAuthError(
                f"Missing user OAuth token. Set {self.env_var} in .env or pass token."
            )
        return value


@dataclass
class AppTokenAuth(ArcGISTokenProvider):
    """Token provider using ArcGIS OAuth client credentials (app authentication)."""

    portal_url: str | None = None
    client_id: str | None = None
    client_secret: str | None = None
    portal_env_var: str = "ARCGIS_PORTAL_URL"
    client_id_env_var: str = "ARCGIS_CLIENT_ID"
    client_secret_env_var: str = "ARCGIS_CLIENT_SECRET"
    _cached_token: str | None = None
    _expires_at_epoch: float = 0.0

    def _require(self, value: str | None, env_var: str) -> str:
        resolved = (value or os.getenv(env_var, "")).strip()
        if not resolved:
            raise ArcGISAuthError(f"Missing required setting: {env_var}")
        return resolved

    def get_token(self, session: requests.Session | None = None) -> str:
        now = time.time()
        if self._cached_token and now < self._expires_at_epoch:
            return self._cached_token

        sess = session or requests.Session()
        portal = self._require(self.portal_url, self.portal_env_var).rstrip("/")
        client_id = self._require(self.client_id, self.client_id_env_var)
        client_secret = self._require(self.client_secret, self.client_secret_env_var)

        token_url = f"{portal}/sharing/rest/oauth2/token"
        payload = {
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "client_credentials",
            "f": "json",
        }

        response = sess.post(token_url, data=payload, timeout=30)
        response.raise_for_status()
        data = response.json()

        if "error" in data:
            err = data["error"]
            message = err.get("error_description") or err.get("message") or str(err)
            raise ArcGISAuthError(f"Failed to acquire app token: {message}")

        token = str(data.get("access_token", "")).strip()
        if not token:
            raise ArcGISAuthError("ArcGIS token endpoint returned no access_token")

        expires_in = int(data.get("expires_in", 3600))
        # Refresh one minute early to avoid race at expiry boundary.
        self._cached_token = token
        self._expires_at_epoch = now + max(0, expires_in - 60)
        return token
