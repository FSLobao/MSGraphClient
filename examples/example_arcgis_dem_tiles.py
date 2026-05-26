"""Fetch DEM tiles from ArcGIS using reusable helpers in arcgisTest package.

Authentication mode is selected via ARCGIS_AUTH_MODE:
    - api_key : ARCGIS_API_KEY
    - app     : ARCGIS_PORTAL_URL + ARCGIS_CLIENT_ID + ARCGIS_CLIENT_SECRET
    - user    : ARCGIS_OAUTH_TOKEN

Usage:
        uv run examples/example_arcgis_dem_tiles.py
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

from arcgisTest.auth import ApiKeyAuth, AppTokenAuth, ArcGISAuthError, UserTokenAuth
from arcgisTest.client import ArcGISImageServiceClient, extract_wms_layer_names

load_dotenv()

REST_URL = "https://mapas.pd.anatel.gov.br/image/rest/services/DEM/ImageServer"
WMS_URL = "https://mapas.pd.anatel.gov.br/image/services/DEM/ImageServer/WMSServer"

# ── Default area of interest (Minas Gerais, Brazil) ──────────────────────────
# Format: lon_min,lat_min,lon_max,lat_max  (EPSG:4326 / WGS-84)
DEFAULT_BBOX = "-40.8,-19.8,-40.7,-19.7"
TILE_SIZE = "512,512"

# ── Output directory ──────────────────────────────────────────────────────────
DOWNLOADS_DIR = Path(__file__).parent.parent / "downloads"


def _build_auth_provider() -> ApiKeyAuth | AppTokenAuth | UserTokenAuth:
    mode = os.getenv("ARCGIS_AUTH_MODE", "api_key").strip().lower()
    if mode == "api_key":
        return ApiKeyAuth()
    if mode == "app":
        return AppTokenAuth()
    if mode == "user":
        return UserTokenAuth()
    raise ArcGISAuthError("Invalid ARCGIS_AUTH_MODE. Use one of: api_key, app, user.")


# ── main ──────────────────────────────────────────────────────────────────────


def main() -> int:
    try:
        auth = _build_auth_provider()
        client = ArcGISImageServiceClient(
            rest_url=REST_URL,
            wms_url=WMS_URL,
            auth=auth,
            timeout=60,
        )
    except (EnvironmentError, ArcGISAuthError) as exc:
        print(f"Configuration error: {exc}")
        return 1

    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)

    # ── 1. ImageServer service info ──────────────────────────────────────────
    print("─" * 60)
    print("1. ImageServer — service info (REST)")
    print("─" * 60)
    try:
        info = client.service_info()
    except Exception as exc:
        print(f"  FAILED: {exc}")
        return 1

    extent = info.get("extent", {})
    sr = info.get("spatialReference", {})
    print(f"  Name         : {info.get('name', '?')}")
    print(f"  Description  : {info.get('description', '(none)')}")
    print(f"  Spatial ref  : WKID {sr.get('wkid', '?')}")
    print(
        f"  Full extent  : xmin={extent.get('xmin')}, ymin={extent.get('ymin')}, "
        f"xmax={extent.get('xmax')}, ymax={extent.get('ymax')}"
    )
    pixel_size = info.get("pixelSizeX") or info.get("meanCellWidth")
    if pixel_size:
        print(f"  Pixel size   : {pixel_size}")

    # ── 2. ImageServer exportImage (TIFF) ────────────────────────────────────
    print()
    print("─" * 60)
    print("2. ImageServer — exportImage (REST → TIFF)")
    print("─" * 60)
    try:
        image_bytes = client.export_image(bbox=DEFAULT_BBOX, size=TILE_SIZE, fmt="tiff")
    except Exception as exc:
        print(f"  FAILED: {exc}")
    else:
        out_path = DOWNLOADS_DIR / "dem_tile.tif"
        out_path.write_bytes(image_bytes)
        print(f"  Saved  : {out_path}")
        print(f"  Size   : {len(image_bytes):,} bytes")

    # ── 3. WMS GetCapabilities ───────────────────────────────────────────────
    print()
    print("─" * 60)
    print("3. WMS — GetCapabilities")
    print("─" * 60)
    wms_layers: list[str] = []
    try:
        capabilities_xml = client.wms_capabilities(version="1.1.1")
    except Exception as exc:
        print(f"  FAILED: {exc}")
    else:
        out_path = DOWNLOADS_DIR / "wms_capabilities.xml"
        out_path.write_text(capabilities_xml, encoding="utf-8")
        print(f"  Saved  : {out_path} ({len(capabilities_xml):,} chars)")
        # Extract layer names for display and use in GetMap below
        wms_layers = extract_wms_layer_names(capabilities_xml)
        if wms_layers:
            print(
                f"  Layers : {', '.join(wms_layers[:10])}"
                + (" …" if len(wms_layers) > 10 else "")
            )

    # ── 4. WMS GetMap (PNG) ──────────────────────────────────────────────────
    print()
    print("─" * 60)
    print("4. WMS — GetMap (PNG)")
    print("─" * 60)
    # Use first real layer from capabilities; fall back to "0" (ArcGIS default)
    layer_name = next((n for n in wms_layers if n not in ("WMS", "WMS_Compat")), "0")
    print(f"  Using layer: {layer_name!r}")
    try:
        wms_tile = client.wms_get_map(
            layers=layer_name,
            bbox=DEFAULT_BBOX,
            size=TILE_SIZE,
            fmt="image/png",
            version="1.1.1",
            srs="EPSG:4326",
        )
    except Exception as exc:
        print(f"  FAILED: {exc}")
    else:
        out_path = DOWNLOADS_DIR / "wms_tile.png"
        out_path.write_bytes(wms_tile)
        print(f"  Saved  : {out_path}")
        print(f"  Size   : {len(wms_tile):,} bytes")

    print()
    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
