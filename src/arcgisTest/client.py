"""Reusable ArcGIS ImageServer + WMS client."""

from __future__ import annotations

import re
from dataclasses import dataclass

import requests

from arcgisTest.auth import ArcGISTokenProvider


class ArcGISServiceError(RuntimeError):
    """Raised when ArcGIS returns a valid response with an error payload."""


def _raise_if_arcgis_error(data: dict) -> None:
    if "error" not in data:
        return
    err = data["error"]
    code = err.get("code", "")
    message = err.get("message", "unknown error")
    raise ArcGISServiceError(f"ArcGIS error {code}: {message}")


def extract_wms_layer_names(capabilities_xml: str) -> list[str]:
    """Extract layer names from a WMS GetCapabilities document."""
    return re.findall(r"<Name>([^<]+)</Name>", capabilities_xml)


@dataclass
class ArcGISImageServiceClient:
    """Client for ArcGIS DEM ImageServer and companion WMS endpoint."""

    rest_url: str
    wms_url: str
    auth: ArcGISTokenProvider
    timeout: int = 30
    session: requests.Session | None = None

    def _session(self) -> requests.Session:
        return self.session or requests.Session()

    def _token(self) -> str:
        return self.auth.get_token(self._session())

    def service_info(self) -> dict:
        response = self._session().get(
            self.rest_url,
            params={"f": "json", "token": self._token()},
            timeout=self.timeout,
        )
        response.raise_for_status()
        data = response.json()
        _raise_if_arcgis_error(data)
        return data

    def export_image(
        self,
        bbox: str,
        size: str = "512,512",
        fmt: str = "tiff",
        bbox_sr: str = "4326",
        image_sr: str = "4326",
    ) -> bytes:
        params = {
            "bbox": bbox,
            "bboxSR": bbox_sr,
            "imageSR": image_sr,
            "size": size,
            "format": fmt,
            "f": "image",
            "token": self._token(),
        }
        response = self._session().get(
            f"{self.rest_url}/exportImage", params=params, timeout=self.timeout
        )
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "").lower()
        if "json" in content_type or "text" in content_type:
            try:
                body = response.json()
                _raise_if_arcgis_error(body)
                rendered = body
            except ValueError:
                rendered = response.text
            raise ArcGISServiceError(
                f"exportImage returned non-image payload: {rendered}"
            )

        return response.content

    def wms_capabilities(self, version: str = "1.1.1") -> str:
        params = {
            "SERVICE": "WMS",
            "REQUEST": "GetCapabilities",
            "VERSION": version,
            "token": self._token(),
        }
        response = self._session().get(
            self.wms_url, params=params, timeout=self.timeout
        )
        response.raise_for_status()
        return response.text

    def wms_get_map(
        self,
        layers: str,
        bbox: str,
        size: str = "512,512",
        fmt: str = "image/png",
        version: str = "1.1.1",
        srs: str = "EPSG:4326",
    ) -> bytes:
        width, height = size.split(",")
        params = {
            "SERVICE": "WMS",
            "REQUEST": "GetMap",
            "VERSION": version,
            "SRS": srs,
            "BBOX": bbox,
            "WIDTH": width,
            "HEIGHT": height,
            "LAYERS": layers,
            "FORMAT": fmt,
            "token": self._token(),
        }
        response = self._session().get(
            self.wms_url, params=params, timeout=self.timeout
        )
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "").lower()
        if "xml" in content_type or "text" in content_type:
            raise ArcGISServiceError(
                f"WMS GetMap returned non-image payload: {response.text[:300]}"
            )

        return response.content
