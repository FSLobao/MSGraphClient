"""MSAL authentication helper for Microsoft Graph.

Supports two authentication modes:
- client_credentials (app-only)
- delegated (user-interactive)

Credentials are received as explicit parameters (environment reading is
handled by :class:`msgraphclient.client.GraphClient`).
"""

import os

import msal

from msgraphclient.client import GraphAuthorizationError, GraphClient  # noqa: F401
from msgraphclient.messages import get_messages
from msgraphclient.settings import (
    GRAPH_DEFAULTS,
    GraphSettings,
    parse_popup_size,
)

# Public API exported by this module.
# GraphAuthorizationError and GraphClient are re-exported from msgraphclient.client.
__all__ = ["GraphAuthorizationError", "GraphClient", "GraphAuthenticator"]


def _token_cache_path() -> str:
    """Return the path for the persistent MSAL delegated token cache file."""
    cache_dir = os.path.join(
        os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
        "MSGraphClient",
    )
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, "token_cache.json")


def _load_token_cache() -> "msal.SerializableTokenCache":
    """Load the MSAL token cache from disk, returning an empty cache on error."""
    cache = msal.SerializableTokenCache()
    path = _token_cache_path()
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except (OSError, ValueError):
            pass
    return cache


def _save_token_cache(cache: "msal.SerializableTokenCache") -> None:
    """Persist the MSAL token cache to disk when its state has changed."""
    if not cache.has_state_changed:
        return
    try:
        with open(_token_cache_path(), "w", encoding="utf-8") as f:
            f.write(cache.serialize())
    except OSError:
        pass


def _find_chromium_app_browser(
    name: str = "_msal_popup",
    popup_size: str = GRAPH_DEFAULTS.auth_popup_size,
) -> str | None:
    """Register a Chromium-based browser in app mode (no address bar or tabs).

    Tries Microsoft Edge then Google Chrome on Windows. When found, registers
    the browser with the ``--app`` flag so the auth page opens in a minimal
    app window instead of a regular tab in an existing browser instance.

    An isolated profile stored under ``%LOCALAPPDATA%\\MSGraphClient\\popup-profile``
    is used so Chromium always applies ``--window-size`` without restoring any
    previously saved window geometry.  The ``--no-signin`` and
    ``--disable-sync`` flags suppress the browser's own account sign-in prompt
    on that profile.  Azure AD session state is managed through MSAL's
    persistent token cache instead, so the browser is only opened on the first
    call (or after a long token expiry).

    Window size is read from the ``GRAPH_AUTH_POPUP_SIZE`` environment variable
    in ``WIDTHxHEIGHT`` format (e.g. ``"600x800"``). Falls back to ``520x680``
    when the variable is absent or has an invalid format.

    Returns the registered browser name to pass as ``browser_name`` to MSAL,
    or ``None`` when no compatible browser is found on the system.
    """
    import webbrowser

    _w, _h = parse_popup_size(popup_size)

    popup_profile = os.path.join(
        os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
        "MSGraphClient",
        "popup-profile",
    )
    os.makedirs(popup_profile, exist_ok=True)

    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            webbrowser.register(
                name,
                None,
                webbrowser.BackgroundBrowser(
                    [
                        path,
                        "--app=%s",
                        f"--window-size={_w},{_h}",
                        f"--user-data-dir={popup_profile}",
                        "--no-first-run",
                        "--no-default-browser-check",
                        "--no-signin",
                        "--disable-sync",
                    ]
                ),
            )
            return name
    return None


class GraphAuthenticator:
    """Authenticate with Azure AD and expose a Graph access token.

    On initialization, supports two modes:
    - default mode: validates env config and acquires a token automatically.
    - injected mode: accepts an explicit ``token`` and/or pre-built ``client``.

    The resolved token is exposed in the public attribute ``token``.

    Public attributes:
        - token
        - auth_mode

    Internal methods (implementation detail):
        - _validate_credentials
        - _acquire_token
        - _acquire_access_token_result

    """

    def __init__(
        self,
        tenant_id: str = "",
        client_id: str = "",
        client_secret: str = "",
        auth_mode: str = "client_credentials",
        redirect_uri: str = "http://localhost",
        delegated_scopes: list[str] | None = None,
        delegated_login_mode: str = "interactive",
        auth_popup_size: str = GRAPH_DEFAULTS.auth_popup_size,
        message_locale: str | None = None,
        token: str | None = None,
        sharepoint_site_id: str = "",
    ) -> None:
        settings = GraphSettings.from_sources(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            sharepoint_site_id=sharepoint_site_id,
            auth_mode=auth_mode,
            redirect_uri=redirect_uri,
            delegated_scopes=delegated_scopes,
            delegated_login_mode=delegated_login_mode,
            auth_popup_size=auth_popup_size,
        )

        self.tenant_id: str = settings.tenant_id
        self.client_id: str = settings.client_id
        self.client_secret: str = settings.client_secret
        self.redirect_uri: str = settings.redirect_uri
        self.delegated_scopes: list[str] = list(settings.delegated_scopes)
        self.delegated_login_mode: str = settings.delegated_login_mode
        self.auth_popup_size: str = settings.auth_popup_size
        self.token: str = ""
        self.sharepoint_site_id: str = settings.sharepoint_site_id
        self.auth_mode = settings.auth_mode
        self.messages = get_messages(message_locale)

        if token:
            self.token = token
        else:
            self._validate_credentials()
            self.token = self._acquire_token()

    @staticmethod
    def _env_file_location() -> str:
        """Return the expected .env file path (cwd-based, as loaded by dotenv)."""
        return os.path.join(os.getcwd(), ".env")

    def _validate_credentials(self) -> None:
        """Validate that required credentials are present for the selected mode.

        Checks both constructor arguments and environment variables.
        Produces a single error message listing all missing items.
        """
        env_path = self._env_file_location()
        docs_url = "https://github.com/FSLobao/MSGraphClient/wiki/Configuration"
        repo_url = "https://github.com/FSLobao/MSGraphClient"

        # Check each required variable (argument OR environment)
        required = [
            (
                "AZURE_TENANT_ID",
                self.tenant_id or os.environ.get("AZURE_TENANT_ID", ""),
            ),
            (
                "AZURE_CLIENT_ID",
                self.client_id or os.environ.get("AZURE_CLIENT_ID", ""),
            ),
            (
                "SHAREPOINT_SITE_ID",
                self.sharepoint_site_id or os.environ.get("SHAREPOINT_SITE_ID", ""),
            ),
        ]

        if self.auth_mode == "client_credentials":
            required.append(
                (
                    "AZURE_CLIENT_SECRET",
                    self.client_secret or os.environ.get("AZURE_CLIENT_SECRET", ""),
                )
            )

        missing = [name for name, value in required if not value]

        if not missing:
            if self.auth_mode == "delegated":
                if self.delegated_login_mode not in ("interactive", "device_code"):
                    raise EnvironmentError(self.messages.invalid_delegated_login_mode)
            return

        raise EnvironmentError(
            f"{self.messages.missing_credentials_header.format(missing=', '.join(missing))}\n"
            f"\n"
            f"{self.messages.missing_credentials_instructions.format(env_path=env_path)}\n"
            f"\n"
            f"{self.messages.missing_credentials_app_only}\n"
            f"\n"
            f"{self.messages.missing_credentials_delegated}\n"
            f"{self.messages.missing_credentials_footer.format(docs_url=docs_url, repo_url=repo_url)}\n"
        )

    def _acquire_token(self) -> str:
        """Acquire and return a Graph API bearer token for the selected mode."""
        if self.auth_mode == "delegated":
            return self._acquire_token_delegated()
        return self._acquire_token_client_credentials()

    def _acquire_token_client_credentials(self) -> str:
        """Acquire and return a Graph API bearer token via client credentials."""
        result = self._acquire_access_token_result(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret,
        )

        if not isinstance(result, dict):
            raise RuntimeError(self.messages.invalid_msal_response)

        if "access_token" not in result:
            error = result.get("error", "unknown")
            description = result.get("error_description", "")
            raise RuntimeError(
                self.messages.token_acquire_failed.format(
                    error=error, description=description
                )
            )

        return result["access_token"]

    def _acquire_token_delegated(self) -> str:
        """Acquire and return a Graph API bearer token via delegated login."""
        result = self._acquire_access_token_result_delegated(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            scopes=self.delegated_scopes,
            login_mode=self.delegated_login_mode,
            redirect_uri=self.redirect_uri,
        )

        if not isinstance(result, dict):
            raise RuntimeError(self.messages.invalid_delegated_msal_response)

        if "access_token" not in result:
            error = result.get("error", "unknown")
            description = result.get("error_description", "")
            raise RuntimeError(
                self.messages.delegated_token_acquire_failed.format(
                    error=error, description=description
                )
            )

        return result["access_token"]

    def _acquire_access_token_result(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
    ) -> dict | None:
        """Acquire token payload from Azure AD via MSAL client credentials flow."""
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )
        result = app.acquire_token_for_client(scopes=GRAPH_DEFAULTS.graph_scopes)
        return result if isinstance(result, dict) else None

    def _acquire_access_token_result_delegated(
        self,
        tenant_id: str,
        client_id: str,
        scopes: list[str],
        login_mode: str,
        redirect_uri: str = "http://localhost",
    ) -> dict | None:
        """Acquire token payload from Azure AD via delegated authentication.

        Tokens are cached in ``%LOCALAPPDATA%\\MSGraphClient\\token_cache.json``
        so the browser is only opened on the first call or after long expiry.
        Subsequent calls are served silently from the cached refresh token.

        MSAL 1.x uses a ``port`` integer parameter (not ``redirect_uri``) for
        acquire_token_interactive.  The port is extracted from ``redirect_uri``
        when it contains one (e.g. "http://localhost:8356" → 8356); otherwise
        MSAL picks a random available port.
        """
        from urllib.parse import urlparse
        import msal.application as _msal_app

        authority = f"https://login.microsoftonline.com/{tenant_id}"
        cache = _load_token_cache()
        app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
            token_cache=cache,
        )

        if login_mode == "device_code":
            # Try silent first; fall back to device-code flow.
            accounts = app.get_accounts()
            if accounts:
                result = app.acquire_token_silent(scopes, account=accounts[0])
                if result and "access_token" in result:
                    _save_token_cache(cache)
                    return result
            flow = app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                return flow if isinstance(flow, dict) else None
            print(flow.get("message", self.messages.device_code_prompt))
            result = app.acquire_token_by_device_flow(flow)
            _save_token_cache(cache)
            return result if isinstance(result, dict) else None

        # Try silent first; fall back to interactive browser.
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                _save_token_cache(cache)
                return result

        parsed = urlparse(redirect_uri)
        port: int | None = parsed.port  # None when no port is specified

        # MSAL already passes browser_name=_preferred_browser() explicitly in
        # acquire_token_interactive, so we cannot inject it via **kwargs (would
        # cause "multiple values" TypeError).  Temporarily replace the internal
        # _preferred_browser function so MSAL picks up our popup browser instead.
        _success_html = (
            "<html><body style='font-family:sans-serif;display:flex;align-items:center;"
            "justify-content:center;height:100vh;margin:0'>"
            f"<p>{self.messages.auth_success_html_text}</p>"
            "<script>window.onload=function(){"
            "window.open('','_self','');window.close();};</script>"
            "</body></html>"
        )
        popup_name = _find_chromium_app_browser(popup_size=self.auth_popup_size)
        if popup_name:
            _orig = _msal_app._preferred_browser
            _msal_app._preferred_browser = lambda: popup_name
            try:
                result = app.acquire_token_interactive(
                    scopes=scopes, port=port, success_template=_success_html
                )
            finally:
                _msal_app._preferred_browser = _orig
        else:
            result = app.acquire_token_interactive(
                scopes=scopes, port=port, success_template=_success_html
            )

        _save_token_cache(cache)
        return result if isinstance(result, dict) else None
