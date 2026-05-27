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
> | Etapa 4 (inscrever site) | Obrigatória | Obrigatória |
>
> Use autenticação delegada quando seus requisitos de governança ou conformidade exigirem trilha de auditoria vinculando cada ação do SharePoint a um usuário nomeado.

Para configuração com credenciais do cliente (sem login de usuário), veja [setup_portal.md](setup_portal.md) ou [setup_cli.md](setup_cli.md).

---

## Quem faz o quê — visão geral

| Etapa | Ação | 🔧 Desenvolvimento | 🔑 Administrador Entra | 🛡️ Administrador SP |
|---|---|:---:|:---:|:---:|
| 1 | Criar registro de aplicativo com URI de redirecionamento e habilitar fluxo de cliente público | ✅ | | |
| 2 | Adicionar `Sites.Selected` como permissão delegada | ✅ | | |
| 3 | Conceder consentimento de administrador para o tenant | | ✅ | |
| 4 | Inscrever o site via `POST /sites/{id}/permissions` | | | ✅ |
| 5 | Configurar `.env` e executar | ✅ | | |

> **Nota:** Neste projeto, a autenticação delegada também é sempre restringida a um site concreto. O usuário autenticado precisa ter acesso ao site e a aplicação precisa ter `Sites.Selected` consentido e uma concessão explícita no site. Não são usadas permissões amplas como `Sites.Read.All` ou `Sites.ReadWrite.All`.

---

## Etapa 1 — Criar o registro de aplicativo com URI de redirecionamento

> 🔧 **Equipe de desenvolvimento**

> ℹ️ **Autenticação delegada não usa segredo de cliente.** O aplicativo age como um cliente público: o usuário faz login diretamente com suas próprias credenciais no Azure AD. Nenhum `AZURE_CLIENT_SECRET` é necessário nem deve ser criado para este fluxo.

1. Abra o [portal do Azure](https://portal.azure.com) → **Microsoft Entra ID** → **Registros de app** → **Novo registro**.
2. **Nome**: escolha um nome descritivo, ex: `MSGraphTest-Delegated`.
3. **Tipos de conta suportados**: selecione **Contas neste diretório organizacional apenas**.
4. Em **URI de Redirecionamento**, selecione o tipo **Aplicativos móveis e de desktop** (não "Web") e informe `http://localhost`.
   > ⚠️ O tipo "Web" causa o erro `AADSTS900971` em runtime. O tipo correto para clientes públicos locais é **Aplicativos móveis e de desktop**.
5. Clique em **Registrar** e anote:
   - **ID do Aplicativo (cliente)** → `AZURE_CLIENT_ID`
   - **ID do Diretório (tenant)** → `AZURE_TENANT_ID`
6. Ainda no registro do aplicativo, vá para **Autenticação**.
7. Role até o final da página, na seção **Configurações avançadas**, localize **Permitir fluxos de cliente público** e selecione **Sim**.
8. Clique em **Salvar**.
   > ⚠️ Sem esta configuração, o Azure AD retorna o erro `AADSTS7000218` ao tentar obter um token interativo, exigindo `client_secret` mesmo sendo um cliente público.

---

## Etapa 2 — Adicionar `Sites.Selected` como permissão delegada

> 🔧 **Equipe de desenvolvimento**

As permissões delegadas controlam o que o aplicativo pode fazer **em nome do usuário autenticado**. Neste projeto, a permissão delegada usada é `Sites.Selected`, para que o aplicativo só consiga operar em sites explicitamente inscritos.

1. No registro de aplicativo, vá para **Permissões de API** → **Adicionar uma permissão** → **Microsoft Graph** → **Permissões delegadas**.
2. Adicione apenas **`Sites.Selected`**.

| Permissão delegada | Necessária para |
|---|---|
| `Sites.Selected` | Restringir o aplicativo aos sites explicitamente inscritos |

> ℹ️ **Não adicione `offline_access`, `openid` ou `profile` manualmente.** O MSAL adiciona esses escopos OIDC automaticamente em todo fluxo interativo. Incluí-los explicitamente na lista de escopos enviada ao Azure AD causa um `ValueError` em runtime.

> `Sites.Selected` não concede acesso a nenhum site por si só. A aplicação só conseguirá ler ou escrever no site depois da inscrição da Etapa 4 e sempre limitada ao que o usuário autenticado já puder fazer nesse mesmo site.

3. Clique em **Adicionar permissões**.

---

## Etapa 3 — Consentimento de administrador

> 🔑 **Administrador Entra**

`Sites.Selected` delegado é classificado pela Microsoft como um escopo de alto privilégio, por isso o Azure AD **bloqueia o auto-consentimento do usuário** independentemente das políticas do tenant. Sem consentimento de administrador, todo usuário que tentar autenticar verá a tela **"Approval required"** e não conseguirá prosseguir.

Para eliminar esta tela para todos os usuários, um Administrador Global ou Administrador de Função com Privilégios deve conceder consentimento uma única vez:

**Opção A — Via Portal do Azure:**
1. No registro de aplicativo, vá para **Permissões de API**.
2. Clique em **Conceder consentimento de administrador para \<tenant\>** e confirme.

**Opção B — Via URL de consentimento de administrador** (substituindo os valores reais):
```
https://login.microsoftonline.com/<AZURE_TENANT_ID>/adminconsent?client_id=<AZURE_CLIENT_ID>
```
Um administrador abre esta URL no navegador e clica em **Aceitar**.

> ⚠️ Esta ação requer **Administrador global** ou **Administrador de função com privilégios**. As funções Administrador de Aplicativos e Administrador de Aplicativos de Nuvem não são suficientes para consentir permissões delegadas do Microsoft Graph.

> ℹ️ Este consentimento não concede acesso a nenhum site. Ele apenas autoriza o uso do escopo `Sites.Selected`; o acesso efetivo aos dados continua dependendo da inscrição do site na Etapa 4.

---

## Etapa 4 — Inscrever o site para o aplicativo

> 🛡️ **Administrador SP** (administrador do tenant do SharePoint)

Siga a mesma inscrição por site descrita na Etapa 4 de [setup_portal.md](setup_portal.md) ou [setup_cli.md](setup_cli.md):

1. Obtenha o `SHAREPOINT_SITE_ID` do site que a aplicação poderá acessar.
2. Execute `POST /sites/{site-id}/permissions` com o `AZURE_CLIENT_ID` da aplicação e o papel `read` ou `write`.
3. Guarde o `site_id` para configuração local e para os scripts de automação em lote.

> O acesso efetivo em runtime será a interseção entre: 1) o papel atribuído ao aplicativo nesse site e 2) as permissões do usuário autenticado no próprio site.

---

## Etapa 5 — Habilitar o modo delegado no pacote

> 🔧 **Equipe de desenvolvimento**

Com a configuração do tenant concluída nas etapas anteriores, configure o modo
delegado via `.env`:

```env
GRAPH_AUTH_MODE=delegated
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_REDIRECT_URI=http://localhost
GRAPH_DELEGATED_LOGIN_MODE=interactive
SHAREPOINT_SITE_ID=contoso.sharepoint.com,<site-guid>,<web-guid>

# Opcional — tamanho da janela de login (padrão: 520x680)
# GRAPH_AUTH_POPUP_SIZE=520x680

# Opcional — escopos delegados adicionais (padrão: Sites.Selected)
# GRAPH_DELEGATED_SCOPES=https://graph.microsoft.com/Sites.Selected
```

> ℹ️ **`AZURE_CLIENT_SECRET` não é usado no modo delegado** e pode ser omitido do `.env`.

> ℹ️ **Não inclua `offline_access`, `openid` ou `profile` em `GRAPH_DELEGATED_SCOPES`.** O MSAL os adiciona automaticamente; incluí-los causa `ValueError`.

**Comportamento da janela de login:**
- Na primeira execução, uma janela de login é aberta no Microsoft Edge ou Chrome em modo aplicativo (sem barra de endereço).
- Após o login bem-sucedido, a janela fecha automaticamente.
- O token de acesso e o refresh token são armazenados em cache em `%LOCALAPPDATA%\MSGraphTest\token_cache.json`. Nas execuções seguintes, o token é renovado silenciosamente sem abrir o navegador (o cache dura enquanto o refresh token for válido, tipicamente ~90 dias de inatividade).
- Para forçar um novo login, delete o arquivo de cache.

Depois, execute um exemplo usando a API class-based:

```bash
uv run examples/example_delegated_site_contents.py
```

Você também pode usar o modo `device_code` (útil em ambientes sem interface gráfica,
como servidores remotos ou containers):

```env
GRAPH_DELEGATED_LOGIN_MODE=device_code
```

Neste modo, o terminal exibe um código e uma URL. O usuário acessa a URL em qualquer
dispositivo, digita o código e faz o login. O script aguarda a conclusão automaticamente.

Em modo delegado, o `GraphClient` e o `GraphAuthenticator` funcionam com a mesma
superfície pública já usada no modo `client_credentials`; muda apenas a forma de
obtenção do token.

Exemplos adicionais ainda válidos:

```bash
uv run examples/example_drive_list.py
uv run examples/example_list_get.py
uv run notebooks/graph_auth_site_attributes.ipynb
```
