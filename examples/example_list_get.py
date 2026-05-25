"""
example_list_get.py — Retrieve and display all items from a SharePoint list.

Usage:
    uv run examples/example_list_get.py
"""

import re

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

from requests import HTTPError

from msgraphtest.graph_client import format_http_error
from msgraphtest.lists import (
    get_list_columns,
    get_list_items,
    get_list_view_columns,
    get_list_views,
)


def _build_column_rename(columns: list[dict]) -> dict[str, str]:
    """Return a mapping of ``fields.<name>`` → ``displayName`` for DataFrame rename."""
    return {
        f"fields.{col['name']}": col["displayName"]
        for col in columns
        if col.get("name") and col.get("displayName")
    }


def _prompt_view_selection(views: list[dict]) -> dict | None:
    """Print numbered view list, prompt the user and return the chosen view dict.

    Returns ``None`` when the user selects 0 (all data).

    Keeps prompting until a valid integer in range [0, len(views)] is entered.
    """
    print("\nViews disponíveis:")
    print("  [0]  Todos os dados da lista")
    for i, view in enumerate(views, start=1):
        print(f"  [{i}]  {view.get('name', view.get('id', '?'))}")

    while True:
        raw = input("\nDigite o número da view desejada (0 para todos): ").strip()
        if raw.isdigit() and 0 <= int(raw) <= len(views):
            choice = int(raw)
            return views[choice - 1] if choice > 0 else None
        print(f"  Entrada inválida. Digite um número entre 0 e {len(views)}.")


def main() -> None:
    """Load SharePoint list items into a DataFrame filtered by a user-selected view.

    Column selection priority:
      1. ``SHAREPOINT_VIEW_COLUMNS`` env var (comma-separated internal field
         names) — bypasses the view API entirely; works with any permission.
      2. Interactive view picker via the Graph API (tries dedicated
         ``/views`` endpoint, then ``?$expand=views`` fallback).
      3. All columns when both above are unavailable.
    """
    import os

    # ── Option C: manual column config from environment ───────────────────────
    env_columns = os.environ.get("SHAREPOINT_VIEW_COLUMNS", "").strip()
    if env_columns:
        internal_names = [c.strip() for c in env_columns.split(",") if c.strip()]
        print(
            f"Usando colunas configuradas via SHAREPOINT_VIEW_COLUMNS: {internal_names}\n"
        )
        # Resolve display names from the full column list (best-effort).
        try:
            all_columns = get_list_columns()
            name_to_display = {
                col["name"]: col["displayName"]
                for col in all_columns
                if col.get("name") and col.get("displayName")
            }
        except HTTPError:
            name_to_display = {}
        columns = [
            {"name": n, "displayName": name_to_display.get(n, n)}
            for n in internal_names
        ]
        view_field_names = {f"fields.{n}" for n in internal_names}
        rename_map = _build_column_rename(columns)
    else:
        # ── Options A/B: interactive view selection via API ───────────────────
        print("Buscando views disponíveis...")
        try:
            views = get_list_views()
        except HTTPError as exc:
            print(
                f"  Aviso: não foi possível obter as views — {format_http_error(exc)}"
            )
            print(
                "  Dica: defina SHAREPOINT_VIEW_COLUMNS no .env para selecionar colunas"
            )
            print("        sem depender da API de views (ex: Title,field_1,field_2).")
            print("  Continuando com todos os dados da lista...\n")
            views = []

        if views:
            selected_view = _prompt_view_selection(views)
        else:
            selected_view = None

        if selected_view:
            view_name = selected_view.get("name", selected_view["id"])
            print(f"\nBuscando colunas da view '{view_name}'...")
            try:
                columns = get_list_view_columns(selected_view["id"])
            except HTTPError as exc:
                print(
                    f"  Aviso: não foi possível obter as colunas da view — {format_http_error(exc)}"
                )
                print("  Usando todas as colunas disponíveis...\n")
                columns = get_list_columns()
            view_field_names = {
                f"fields.{col['name']}" for col in columns if col.get("name")
            }
        else:
            print("\nBuscando todas as definições de coluna...")
            columns = get_list_columns()
            view_field_names = None

        rename_map = _build_column_rename(columns)

    print("Buscando itens da lista do SharePoint...\n")
    items = get_list_items()
    if not items:
        print("(nenhum item encontrado)")
        return

    df = pd.json_normalize(items)
    if view_field_names is not None:
        field_cols = [c for c in df.columns if c in view_field_names]
    else:
        field_cols = [c for c in df.columns if c.startswith("fields.")]

    # Build a content DataFrame with only Title + field_0..99 columns.
    content_field_cols = [
        c
        for c in field_cols
        if c == "fields.Title" or re.fullmatch(r"fields\.field_(?:\d|[1-9]\d)", c)
    ]
    content_rename_map = {
        c: rename_map.get(c, c.removeprefix("fields.")) for c in content_field_cols
    }
    df_list_content = df[content_field_cols].rename(columns=content_rename_map).copy()

    # Everything else is treated as SharePoint metadata.
    metadata_cols = [c for c in df.columns if c not in content_field_cols]
    df_sharepoint_metadata = df[metadata_cols].copy()

    print("Conteudo da Lista")
    print(df_list_content.head())

    print("\nSharePoint Metadata")
    print(df_sharepoint_metadata.head())

    print(f"\nTotal de linhas carregadas: {len(df)}")


if __name__ == "__main__":
    main()
