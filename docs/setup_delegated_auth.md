# Configuração passo a passo para autenticação delegada (usuário)

Este guia configura o acesso ao SharePoint via Microsoft Graph API usando o **fluxo de código de autorização OAuth 2.0** (autenticação delegada), onde o aplicativo age em nome de um usuário autenticado.

> **Diferença em relação à autenticação por credenciais do cliente:**
>
> | | Credenciais do cliente (app-only) | Autenticação delegada (usuário) |
> |---|---|---|
> | Quem executa | O próprio aplicativo | Um usuário autenticado |
> | Login necessário | Não — execução sem supervisão | Sim — a cada execução |
> | Trilha de auditoria | Identidade do aplicativo | Identidade do usuário individual |
> | Permissões necessárias | Permissões de aplicativo (`Role`) | Permissões delegadas (`Scope`) |
> | Etapa 4 (inscrever site) | Obrigatória | **Não necessária** — o acesso do usuário ao site é respeitado diretamente |
>
> Use autenticação delegada quando seus requisitos de governança ou conformidade exigirem trilha de auditoria vinculando cada ação do SharePoint a um usuário nomeado.

Para configuração com credenciais do cliente (sem login de usuário), veja [setup_portal.md](setup_portal.md) ou [setup_cli.md](setup_cli.md).

---

## Quem faz o quê — visão geral

| Etapa | Ação | 🔧 Desenvolvimento | 🔑 Administrador Entra |
|---|---|:---:|:---:|
| 1 | Criar registro de aplicativo com URI de redirecionamento | ✅ | |
| 2 | Adicionar permissões delegadas do SharePoint | ✅ | |
| 3 | Conceder consentimento de administrador (opcional) | | ✅ |
| 4 | Adaptar `auth.py` para o fluxo interativo | ✅ | |
| 5 | Configurar `.env` e executar | ✅ | |

> **Nota:** A Etapa 4 do guia de credenciais do cliente (inscrever o site via `POST /sites/{id}/permissions`) **não existe** neste fluxo. O acesso do aplicativo é limitado ao que o usuário autenticado já tem permissão no SharePoint — nenhuma configuração adicional de administrador SP é necessária.

---

## Etapa 1 — Criar o registro de aplicativo com URI de redirecionamento

> 🔧 **Equipe de desenvolvimento**

1. Abra o [portal do Azure](https://portal.azure.com) → **Microsoft Entra ID** → **Registros de app** → **Novo registro**.
2. **Nome**: escolha um nome descritivo, ex: `MSGraphTest-Delegated`.
3. **Tipos de conta suportados**: selecione **Contas neste diretório organizacional apenas**.
4. Em **URI de Redirecionamento**, selecione o tipo **Web** e informe `http://localhost:8000` (para execução local).
5. Clique em **Registrar** e anote:
   - **ID do Aplicativo (cliente)** → `AZURE_CLIENT_ID`
   - **ID do Diretório (tenant)** → `AZURE_TENANT_ID`
6. Vá para **Certificados e segredos** → **Segredos do cliente** → **Novo segredo do cliente**.
7. Defina uma descrição e uma expiração alinhada com sua política de rotação (máximo 24 meses).
8. Clique em **Adicionar** e **copie imediatamente** o valor do segredo → `AZURE_CLIENT_SECRET`.

> ⚠️ O valor do segredo é mostrado apenas uma vez. Armazene-o com segurança. Se você navegar para longe sem copiá-lo, delete e recrie-o.

---

## Etapa 2 — Adicionar permissões delegadas do SharePoint

> 🔧 **Equipe de desenvolvimento**

As permissões delegadas controlam o que o aplicativo pode fazer **em nome do usuário autenticado**. O aplicativo nunca terá mais acesso do que o próprio usuário possui no SharePoint.

1. No registro de aplicativo, vá para **Permissões de API** → **Adicionar uma permissão** → **Microsoft Graph** → **Permissões delegadas**.
2. Adicione as permissões necessárias para o seu caso de uso:

| Permissão delegada | Necessária para |
|---|---|
| `Sites.Read.All` | Listar sites, drives e arquivos; ler listas |
| `Sites.ReadWrite.All` | Criar/atualizar itens de lista; enviar arquivos |
| `Files.Read.All` | Baixar conteúdo de arquivos em drives |
| `Files.ReadWrite.All` | Enviar ou sobrescrever arquivos em drives |
| `offline_access` | Obter refresh token para renovar a sessão sem novo login |

> Use o conjunto mínimo que sua aplicação precisa. Para casos de somente leitura, `Sites.Read.All` e `Files.Read.All` são suficientes.

3. Clique em **Adicionar permissões**.

---

## Etapa 3 — Consentimento de administrador (opcional)

> 🔑 **Administrador Entra** (opcional)

Por padrão, cada usuário deverá consentir individualmente com as permissões na primeira execução. Para suprimir essa tela de consentimento para todos os usuários do tenant:

1. No registro de aplicativo, vá para **Permissões de API**.
2. Clique em **Conceder consentimento de administrador para \<tenant\>** e confirme.

> Se este passo for omitido, cada usuário verá uma tela de consentimento na primeira vez que executar o aplicativo — o que é aceitável na maioria dos cenários.

---

## Etapa 4 — Adaptar `auth.py` para o fluxo interativo

> 🔧 **Equipe de desenvolvimento**

Substitua o conteúdo de `src/msgraphtest/auth.py` pelo seguinte:

```python
"""
auth.py — MSAL delegated-auth (authorization code flow) helper.

Acquires an access token for the Microsoft Graph API using the
OAuth 2.0 authorization code flow (delegated / user identity).
On first run, opens a browser for the user to sign in. Subsequent
runs reuse the cached token silently until it expires.

Required environment variables:
    AZURE_TENANT_ID     – Azure AD tenant ID
    AZURE_CLIENT_ID     – App registration client ID
    AZURE_CLIENT_SECRET – App registration client secret
    AZURE_REDIRECT_URI  – Redirect URI registered in the app (default: http://localhost:8000)

Usage::

    from msgraphtest.auth import get_access_token
    token = get_access_token()
"""

import os

import msal
from dotenv import load_dotenv

load_dotenv()

GRAPH_SCOPES = [
    "https://graph.microsoft.com/Sites.ReadWrite.All",
    "https://graph.microsoft.com/Files.ReadWrite.All",
    "offline_access",
]

_TOKEN_CACHE_FILE = ".msal_token_cache.json"


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(_TOKEN_CACHE_FILE):
        with open(_TOKEN_CACHE_FILE) as f:
            cache.deserialize(f.read())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        with open(_TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def _get_config() -> tuple[str, str, str, str]:
    tenant_id = os.environ.get("AZURE_TENANT_ID", "")
    client_id = os.environ.get("AZURE_CLIENT_ID", "")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET", "")
    redirect_uri = os.environ.get("AZURE_REDIRECT_URI", "http://localhost:8000")
    if not all([tenant_id, client_id, client_secret]):
        raise EnvironmentError(
            "Missing one or more required environment variables: "
            "AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET"
        )
    return tenant_id, client_id, client_secret, redirect_uri


def get_access_token() -> str:
    """Return a valid Graph API bearer token using delegated (user) auth.

    Attempts silent token acquisition from cache first. If no cached token
    is available or it has expired, opens a browser for the user to sign in
    interactively.
    """
    tenant_id, client_id, client_secret, redirect_uri = _get_config()
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    cache = _load_cache()

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
        token_cache=cache,
    )

    # Try silent acquisition from cache
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])

    # Fall back to interactive login
    if not result:
        result = app.acquire_token_interactive(
            scopes=GRAPH_SCOPES,
            redirect_uri=redirect_uri,
        )

    _save_cache(cache)

    if "access_token" not in result:
        error = result.get("error", "unknown")
        description = result.get("error_description", "")
        raise RuntimeError(f"Failed to acquire token: {error} — {description}")

    return result["access_token"]
```

> **Segurança:** O arquivo `.msal_token_cache.json` contém tokens de acesso e de atualização. Adicione-o ao `.gitignore` para não versioná-lo.

Adicione ao `.gitignore`:

```
.msal_token_cache.json
```

---

## Etapa 5 — Configurar `.env` e executar

> 🔧 **Equipe de desenvolvimento**

```bash
uv sync
cp .env.example .env
```

Edite `.env` com os valores coletados. A variável `SHAREPOINT_SITE_ID` ainda é necessária; `AZURE_REDIRECT_URI` é opcional se usar o valor padrão `http://localhost:8000`:

```ini
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=<your-secret-value>
AZURE_REDIRECT_URI=http://localhost:8000
SHAREPOINT_SITE_ID=contoso.sharepoint.com,<site-guid>,<web-guid>
SHAREPOINT_DRIVE_ID=b!<drive-id>
SHAREPOINT_LIST_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

> Para descobrir `SHAREPOINT_SITE_ID`, `SHAREPOINT_DRIVE_ID` e `SHAREPOINT_LIST_ID`, use o [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) autenticado com sua conta de usuário, conforme descrito nas Etapas 3 e 5 de [setup_portal.md](setup_portal.md).

Na primeira execução, um navegador será aberto para login. Após autenticação, o token é armazenado em cache e as execuções seguintes serão silenciosas:

```bash
uv run examples/example_drive_list.py
uv run examples/example_list_get.py
```
