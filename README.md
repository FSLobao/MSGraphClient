# MSGraphTest — SharePoint via Microsoft Graph API

Um projeto de teste em Python demonstrando como acessar o SharePoint através da
**Microsoft Graph API** usando MSAL para autenticação. As operações abordadas
incluem gerenciamento de biblioteca de documentos (drive) e manipulação de listas do SharePoint.

Licenciado sob a [GNU General Public License v3.0](LICENSE).

---

## Estrutura do projeto

```
MSGraphTest/
├── src/
│   └── msgraphtest/
│       ├── __init__.py        # ponto de entrada do pacote
│       ├── auth.py            # auxiliar de token client-credentials (credenciais do cliente) com MSAL
│       ├── graph_client.py    # wrapper (invólucro) HTTP fino para chamadas REST do Graph
│       ├── drive.py           # operações de biblioteca de documentos
│       └── lists.py           # operações de lista do SharePoint
├── tests/
│   ├── test_auth.py
│   ├── test_drive.py
│   └── test_lists.py
├── examples/
│   ├── example_drive_list.py       # listar conteúdo raiz da drive
│   ├── example_drive_download.py   # baixar arquivo para pasta local
│   ├── example_drive_upload.py     # enviar arquivo local
│   ├── example_drive_read_write.py # ler e atualizar conteúdo de texto do arquivo
│   ├── example_list_get.py         # recuperar todos os itens de lista
│   ├── example_list_create.py      # criar item de lista
│   └── example_list_update.py      # atualizar item de lista
├── docs/
│   └── getting_started.md
├── downloads/                 # (ignorado por git) destino de download local
├── .env.example               # copie para .env e preencha as credenciais
├── pyproject.toml
└── LICENSE
```

---

## Pré-requisitos

| Requisito | Observações |
|---|---|
| Python ≥ 3.11 | Testado com 3.11+ |
| [UV](https://docs.astral.sh/uv/) | Gerenciador de pacotes e ambiente virtual |
| Registro de Aplicativo do Azure AD (inscrição de app) | Com permissões MS Graph `Sites.Read.All` / `Sites.ReadWrite.All` |

---

## Início rápido

### 1. Clonar e instalar dependências

```bash
git clone <repo-url>
cd MSGraphTest
uv sync
```

### 2. Configurar credenciais

```bash
cp .env.example .env
# edite .env com seus detalhes do Azure AD e SharePoint
```

Variáveis obrigatórias em `.env`:

| Variável | Descrição |
|---|---|
| `AZURE_TENANT_ID` | ID do tenant (locatário) do Azure AD |
| `AZURE_CLIENT_ID` | ID do cliente do registro de aplicativo |
| `AZURE_CLIENT_SECRET` | Segredo do cliente do registro de aplicativo |
| `SHAREPOINT_SITE_ID` | ID do site do Graph (ex: `contoso.sharepoint.com,guid,guid`) |
| `SHAREPOINT_DRIVE_ID` | ID da drive (unidade) da biblioteca de documentos |
| `SHAREPOINT_LIST_ID` | ID da lista para operações de lista |

> **Encontrando IDs** — veja [docs/getting_started.md](docs/getting_started.md).

### 3. Executar um exemplo

```bash
uv run examples/example_drive_list.py
uv run examples/example_list_get.py
```

---

## Executando testes

```bash
uv run pytest
```

Relatório de cobertura é impresso automaticamente. Os testes usam mocking e **não**
requerem credenciais reais.

---

## Module overview

### `auth.py`
Acquires a Graph API bearer token using the OAuth 2.0 **client credentials**
flow via [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python).

### `graph_client.py`
`GraphClient` — a thin `requests.Session` wrapper that injects the bearer
token and exposes `get`, `post`, `patch`, `put_bytes`, and `get_raw` helpers.

### `drive.py`
Document library operations:

| Function | Description |
|---|---|
| `list_drive_items(folder_path)` | List children of a folder |
| `download_file(item_id, local_path)` | Download a file to disk |
| `upload_file(local_path, remote_folder)` | Upload a local file (≤ 4 MB) |
| `read_file_content(item_id)` | Return file text as a string |
| `write_file_content(item_id, content)` | Overwrite a file's text content |

### `lists.py`
SharePoint list operations:

| Function | Description |
|---|---|
| `get_list_items(select)` | Retrieve all items (optionally select fields) |
| `create_list_item(fields)` | Create a new item |
| `update_list_item(item_id, fields)` | Update fields on an existing item |

---

## License

This project is licensed under the **GNU General Public License v3.0**.
See [LICENSE](LICENSE) for the full text.
