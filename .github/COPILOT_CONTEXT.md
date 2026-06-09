# Copilot Context — ezspi

Este arquivo resume o estado atual do repositório para manutenção de código e documentação.
Atualize-o sempre que a superfície pública, dependências ou estrutura do projeto mudarem.

---

## Visão Geral

**Objetivo**: biblioteca Python para abstrair autenticação com MSAL e simplificar acesso ao SharePoint Online via Microsoft Graph API, com foco em privilégio mínimo (`Sites.Selected`).

**Fluxos suportados**:
- `client_credentials` (app-only)
- `delegated` (`interactive` ou `device_code`)

**Componentes públicos principais**:
- `Authenticator`: aquisição de token e validação de credenciais/configuração.
- `Client`: sessão HTTP para Graph, helpers de request e normalização de erros HTTP.
- `SPLibrary`: operações de biblioteca de documentos (navegação, upload/download, leitura/escrita).
- `SPList`: operações de listas (schema, views, leitura paginada, validação e save).

No fluxo `delegated`, o acesso efetivo em runtime continua sendo a interseção entre:
- concessão do aplicativo no site (`Sites.Selected`)
- permissões reais do usuário autenticado nesse mesmo site.

---

## Stack

- Python `>=3.11` (README indica testes com 3.14)
- `uv` (ambiente/dependências)
- `pytest`, `pytest-cov`, `pytest-mock`
- `requests`
- `msal`
- `python-dotenv`
- `pandas`
- `python-dateutil`

---

## Estrutura Atual

```text
ezspi/
├── docs/
│   ├── bulk_create_apps.md
│   ├── getting_started.md
│   ├── setup_cli.md
│   ├── setup_delegated_auth.md
│   └── setup_portal.md
├── examples/
│   ├── example_drive_download.py
│   ├── example_drive_folder_operations.py
│   ├── example_drive_list.py
│   ├── example_drive_read_write.py
│   ├── example_drive_upload.py
│   ├── example_list_create.py
│   ├── example_list_get.py
│   ├── example_list_update.py
│   ├── example_site_contents.py
│   ├── list_value_generation.py
│   └── downloads/
├── notebooks/
│   ├── graph_auth_site_attributes.ipynb
│   └── downloads/
├── src/
│   ├── bulkCreate/
│   └── ezspi/
│       ├── __init__.py
│       ├── auth.py
│       ├── client.py
│       ├── drive.py
│       ├── lists.py
│       ├── messages.py
│       ├── settings.py
│       └── locales/
├── tests/
│   ├── test_auth.py
│   ├── test_drive.py
│   ├── test_graph_client.py
│   ├── test_lists.py
│   ├── test_list_value_generation.py
│   ├── test_settings.py
│   └── test_site.py
├── pyproject.toml
└── README.md
```

---

## Superfície de API (Atual)

### `src/ezspi/client.py`
- `AuthorizationError` (especializa falhas 401/403).
- `Client.format_http_error(error)`
- `Client.get(path, **kwargs)`
- `Client.post(path, json, **kwargs)`
- `Client.patch(path, json, **kwargs)`
- `Client.put_bytes(path, data, content_type=..., **kwargs)`
- `Client.get_raw(path, **kwargs)`
- `Client.get_raw_with_encoding(path, **kwargs)`

Também popula atributos de site (`site_graph_id`, `site_name`, `site_display_name`, `site_web_url`, `site_drives`, `site_lists`) quando `SHAREPOINT_SITE_ID` está disponível.

### `src/ezspi/auth.py`
- `Authenticator` com resolução por `Settings` e suporte a `client_credentials` e `delegated`.
- Reexporta `Client` e `AuthorizationError`.
- Usa cache de token em memória para fluxo delegado.

### `src/ezspi/drive.py`
- `SPLibrary.pwd()`
- `SPLibrary.cd(path)`
- `SPLibrary.ls(path=None)`
- `SPLibrary.download(item_id, local_path)`
- `SPLibrary.upload(local_path, remote_folder="root", remote_name=None)`
- `SPLibrary.read(item_id, encoding=None)`
- `SPLibrary.write(item_id, content, encoding=None)`

`read()` detecta charset da resposta HTTP quando possível e persiste em `last_encoding`, usado por `write()` quando `encoding` não é informado.

### `src/ezspi/lists.py`
- `SPList.get_views()` (com fallback seguro)
- `SPList.get_view_columns(view_id)`
- `SPList.get_columns(names=None)`
- `SPList.get_schema()`
- `SPList.get_field_types()`
- `SPList.validate_item(data)`
- `SPList.get_items(select=None, include_id=True)`
- `SPList.get_item_template(include_optional=True)`
- `SPList.get_items_dataframe(select=None, include_id=True)`
- `SPList.save_dataframe(dataframe)`
- `SPList.save_item(data)`
- `SPList.save_items(items)`

---

## Documentação e Exemplos

- README concentra onboarding, autenticação (`client_credentials` e `delegated`), execução de exemplos e testes.
- Notebook principal de validação end-to-end: `notebooks/graph_auth_site_attributes.ipynb`.
- Guias operacionais estão em `docs/`, incluindo setup por CLI/Portal e criação em lote de apps.

---

## Testes e Validação

Comando recomendado:

```bash
uv run pytest tests/
```

Observação desta revisão:
- A execução local não foi concluída nesta sessão devido a erro de ambiente ao invocar `uv run` (`Failed to canonicalize script path`).
- Não afirmar contagem de testes aprovados sem nova execução local ou CI.

---

## Segurança

Nunca versionar:
- `.env`
- caches de token
- outputs contendo segredos/tokens

Diretrizes:
- manter privilégio mínimo (`Sites.Selected`)
- restringir grants por site
- usar segredos com rotação/expiração explícita

---

## Última Atualização

- Data: 09/06/2026
- Alteração principal: contexto sincronizado com estrutura atual (`src/ezspi`, exemplos e testes), superfície pública revisada e seção de validação ajustada para refletir o estado real da sessão.

