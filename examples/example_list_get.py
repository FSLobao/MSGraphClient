"""
example_list_get.py — Retrieve and display all items from a SharePoint list.

Usage:
    uv run examples/example_list_get.py
"""

import pandas as pd
from dotenv import load_dotenv
from requests import HTTPError

from msgraphtest.auth import GraphClient
from msgraphtest.lists import GraphList

load_dotenv()


def _build_column_rename(columns: list[dict]) -> dict[str, str]:
    """Return a mapping of internal column names to display names."""
    return {
        col["name"]: col["displayName"]
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

        client = GraphClient()
        list_client = GraphList(client=client)

    # ── Option C: manual column config from environment ───────────────────────
    env_columns = os.environ.get("SHAREPOINT_VIEW_COLUMNS", "").strip()
    if env_columns:
        internal_names = [c.strip() for c in env_columns.split(",") if c.strip()]
        print(
            f"Usando colunas configuradas via SHAREPOINT_VIEW_COLUMNS: {internal_names}\n"
        )
        try:
            filtered_columns = list_client.get_list_columns(
                names=["Title", *internal_names]
            )
        except HTTPError as exc:
            print(
                "  Aviso: não foi possível resolver display names das colunas "
                f"designadas — {GraphClient.format_http_error(exc)}"
            )
            filtered_columns = []

        display_name_map = {
            col["name"]: col["displayName"]
            for col in filtered_columns
            if col.get("name") and col.get("displayName")
        }
        columns = [
            {"name": n, "displayName": display_name_map.get(n, n)}
            for n in ["Title", *internal_names]
        ]
        rename_map = _build_column_rename(columns)
    else:
        # ── Options A/B: interactive view selection via API ───────────────────
        print("Buscando views disponíveis...")
        try:
            views = list_client.get_list_views()
        except HTTPError as exc:
            print(
                f"  Aviso: não foi possível obter as views — {GraphClient.format_http_error(exc)}"
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
                columns = list_client.get_list_view_columns(selected_view["id"])
            except HTTPError as exc:
                print(
                    f"  Aviso: não foi possível obter as colunas da view — {GraphClient.format_http_error(exc)}"
                )
                print("  Usando todas as colunas disponíveis...\n")
                columns = list_client.get_list_columns()
        else:
            print("\nBuscando todas as definições de coluna...")
            columns = list_client.get_list_columns()

        rename_map = _build_column_rename(columns)

    print("Buscando itens da lista do SharePoint...\n")
    selected_fields = [
        col["name"] for col in columns if col.get("name") and col["name"] != "Title"
    ]
    items = list_client.get_list_items(
        select=selected_fields,
        include_title=True,
        fields_only=True,
        include_item_id=True,
    )
    if not items:
        print("(nenhum item encontrado)")
        return

    df_list_content = pd.DataFrame(items).rename(columns=rename_map).copy()

    print("Conteudo da Lista")
    print(df_list_content.head())
    print(f"\nTotal de linhas carregadas: {len(df_list_content)}")


if __name__ == "__main__":
    main()
