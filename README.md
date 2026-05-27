<a id="topo"></a>

# MSGraphTest — SharePoint via Microsoft Graph API

Um projeto de teste em Python demonstrando como acessar o SharePoint através da
**Microsoft Graph API** usando MSAL para autenticação. As operações abordadas
incluem gerenciamento de biblioteca de documentos (drive) e manipulação de listas do SharePoint.

O projeto adota **privilégio mínimo** como regra: o acesso ao SharePoint é feito com
`Sites.Selected`, sempre restrito a sites explicitamente inscritos. Isso vale tanto
para autenticação `app_only` quanto para autenticação `delegated`.

Licenciado sob a [GNU General Public License v3.0](LICENSE).

---

<details>
	<summary><strong>Índice</strong></summary>
	<ul>
		<li><a href="#sec-estrutura-do-projeto">Estrutura do projeto</a></li>
		<li><a href="#sec-pre-requisitos">Pre-requisitos</a></li>
		<li><a href="#sec-inicio-rapido">Inicio rapido</a></li>
		<li><a href="#sec-notebooks-interativos">Notebooks interativos</a></li>
		<li><a href="#sec-executando-testes">Executando testes</a></li>
		<li><a href="#sec-criacao-em-lote">Criacao em lote de aplicacoes Azure AD</a></li>
		<li><a href="#sec-visao-geral-modulos">Visao geral dos modulos</a></li>
		<li><a href="#sec-documentacao-adicional">Documentacao adicional</a></li>
		<li><a href="#sec-licenca">Licenca</a></li>
	</ul>
</details>

---

<a id="sec-estrutura-do-projeto"></a>

## Estrutura do projeto

Escolha a visualização que preferir. Esta seção traz 3 alternativas para comparar renderização e legibilidade.

### Opção 1: Árvore textual (formato atual)

```
MSGraphTest/
├── src/
│   ├── bulkCreate/
│   │   ├── bulk_create_apps.py        # utilitário para criar múltiplas apps em lote (Python)
│   │   └── Bulk-CreateApps.ps1        # idem, versão PowerShell
│   └── msgraphtest/
│       ├── __init__.py                # ponto de entrada do pacote
│       ├── auth.py                    # GraphClient + GraphAuthenticator
│       ├── drive.py                   # operações de biblioteca de documentos
│       └── lists.py                   # operações de lista do SharePoint
├── tests/
│   ├── test_auth.py
│   ├── test_drive.py
│   └── test_lists.py
├── examples/
│   ├── example_drive_list.py          # listar conteúdo raiz da drive
│   ├── example_drive_download.py      # baixar arquivo para pasta local
│   ├── example_drive_upload.py        # enviar arquivo local
│   ├── example_drive_read_write.py    # ler e atualizar conteúdo de texto do arquivo
│   ├── example_list_get.py            # recuperar todos os itens de lista
│   ├── example_list_create.py         # criar item de lista
│   ├── example_list_update.py         # atualizar item de lista
│   └── bulk_create_example.json       # modelo de entrada para bulk_create_apps
├── notebooks/
│   └── graph_auth_site_attributes.ipynb  # fluxo interativo end-to-end (auth + drive + lists)
├── docs/
│   ├── getting_started.md             # guia de início rápido
│   ├── setup_cli.md                   # setup com Azure CLI / PowerShell
│   ├── setup_portal.md                # setup com Azure Portal
│   ├── setup_delegated_auth.md        # setup com autenticação delegada (usuário)
│   └── bulk_create_apps.md            # documentação de criação em lote de apps
├── downloads/                 # (ignorado por git) destino de download local
├── .env.example               # copie para .env e preencha as credenciais
├── pyproject.toml
└── LICENSE
```

### Opção 2: Tabela estruturada (nome e descrição separados)

| Caminho | Tipo | Descrição |
|---|---|---|
| src/bulkCreate/bulk_create_apps.py | Script Python | Utilitário para criar múltiplas apps em lote |
| src/bulkCreate/Bulk-CreateApps.ps1 | Script PowerShell | Utilitário equivalente em PowerShell |
| src/msgraphtest/__init__.py | Módulo | Ponto de entrada do pacote |
| src/msgraphtest/auth.py | Módulo | GraphClient + GraphAuthenticator |
| src/msgraphtest/drive.py | Módulo | Operações de biblioteca de documentos |
| src/msgraphtest/lists.py | Módulo | Operações de lista do SharePoint |
| tests/ | Diretório | Testes automatizados |
| examples/ | Diretório | Exemplos executáveis de uso |
| notebooks/graph_auth_site_attributes.ipynb | Notebook | Fluxo interativo end-to-end (auth + drive + lists) |
| docs/ | Diretório | Guias e documentação de setup |
| downloads/ | Diretório | Destino de downloads locais (ignorado por git) |
| .env.example | Arquivo de configuração | Modelo para variáveis de ambiente |
| pyproject.toml | Arquivo de projeto | Dependências e configuração Python |
| LICENSE | Licença | GPL v3.0 |

### Opção 3: Estrutura HTML colapsável

<details>
	<summary><strong>Expandir estrutura por grupos</strong></summary>

	<details>
		<summary><strong>src/</strong> - codigo-fonte principal</summary>

		<details>
			<summary><strong>bulkCreate/</strong> - criacao de apps em lote</summary>
			<ul>
				<li><code>bulk_create_apps.py</code> - criacao em lote (Python)</li>
				<li><code>Bulk-CreateApps.ps1</code> - criacao em lote (PowerShell)</li>
			</ul>
		</details>

		<details>
			<summary><strong>msgraphtest/</strong> - modulos de Graph e SharePoint</summary>
			<ul>
				<li><code>__init__.py</code> - ponto de entrada do pacote</li>
				<li><code>auth.py</code> - cliente e autenticador Graph</li>
				<li><code>drive.py</code> - operacoes de drive</li>
				<li><code>lists.py</code> - operacoes de listas</li>
			</ul>
		</details>
	</details>

	<details>
		<summary><strong>tests/</strong> - testes automatizados</summary>
		<ul>
			<li><code>test_auth.py</code></li>
			<li><code>test_drive.py</code></li>
			<li><code>test_lists.py</code></li>
		</ul>
	</details>

	<details>
		<summary><strong>examples/</strong> - scripts de exemplo</summary>
		<ul>
			<li><code>example_drive_list.py</code></li>
			<li><code>example_drive_download.py</code></li>
			<li><code>example_drive_upload.py</code></li>
			<li><code>example_drive_read_write.py</code></li>
			<li><code>example_list_get.py</code></li>
			<li><code>example_list_create.py</code></li>
			<li><code>example_list_update.py</code></li>
			<li><code>bulk_create_example.json</code></li>
		</ul>
	</details>

	<details>
		<summary><strong>notebooks/</strong> - testes interativos</summary>
		<ul>
			<li><code>graph_auth_site_attributes.ipynb</code></li>
		</ul>
	</details>

	<details>
		<summary><strong>docs/</strong> - documentacao adicional</summary>
		<ul>
			<li><code>getting_started.md</code></li>
			<li><code>setup_cli.md</code></li>
			<li><code>setup_portal.md</code></li>
			<li><code>setup_delegated_auth.md</code></li>
			<li><code>bulk_create_apps.md</code></li>
		</ul>
	</details>

	<details>
		<summary><strong>downloads/</strong> - artefatos locais (gitignored)</summary>
		<ul>
			<li>Arquivos de download e upload de teste gerados localmente</li>
		</ul>
	</details>

	<ul>
		<li><code>.env.example</code> - modelo de configuracao</li>
		<li><code>pyproject.toml</code> - configuracao do projeto</li>
		<li><code>LICENSE</code> - licenca do repositorio</li>
	</ul>
</details>

[⬆ Voltar ao topo](#topo)

---

<a id="sec-pre-requisitos"></a>

## Pré-requisitos

| Requisito | Observações |
|---|---|
| Python ≥ 3.11 | Testado com 3.11+ |
| [UV](https://docs.astral.sh/uv/) | Gerenciador de pacotes e ambiente virtual |
| Registro de aplicativo no Microsoft Entra ID | Configure `Sites.Selected` e inscreva os sites necessários |

> Este repositório **não usa** permissões amplas como `Sites.Read.All` ou `Sites.ReadWrite.All` para acesso a dados no SharePoint.

[⬆ Voltar ao topo](#topo)

---

<a id="sec-inicio-rapido"></a>

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

### 3. Escolher o modelo de autenticação

- **`app_only`**: indicado para automação sem interação do usuário.
- **`delegated`**: indicado quando é necessário associar as ações a um usuário autenticado.

Nos dois casos, o projeto usa `Sites.Selected` e exige inscrição explícita do site.
No fluxo `delegated`, o acesso efetivo é a interseção entre a concessão do aplicativo
no site e as permissões que o usuário já possui nesse mesmo site.

### 4. Executar um exemplo

```bash
uv run examples/example_drive_list.py
uv run examples/example_list_get.py
```

[⬆ Voltar ao topo](#topo)

---

<a id="sec-notebooks-interativos"></a>

## Notebooks interativos (alternativa aos examples)

Além dos scripts em `examples/`, você pode usar notebooks para validar o fluxo de forma
interativa, inspecionando respostas e DataFrames a cada etapa.

Notebook principal:

- `notebooks/graph_auth_site_attributes.ipynb`

Esse notebook executa um fluxo end-to-end para:

1. carregar `.env` e autenticar no Graph;
2. consultar atributos e conteúdo do site (drives/lists);
3. testar operações de conteúdo no drive (write, update, load e download);
4. testar criação e atualização de itens em lista com visualização tabular.

Quando usar notebooks em vez dos examples:

- quando você quer depurar autenticação passo a passo;
- quando precisa validar transformação e inspeção de dados em DataFrames;
- quando deseja testar rapidamente mudanças no processo de edição de conteúdo.

> [!WARNING]
> **Boas práticas de higiene (credenciais e dados sensíveis)**
>
> - Nunca imprima valores de variáveis sensíveis do `.env` (como `AZURE_CLIENT_SECRET`).
> - Evite exibir tokens, headers de autorização ou payloads contendo segredos.
> - Limpe os outputs do notebook antes de commit (`Clear All Outputs`) para não versionar dados sensíveis.
> - Mantenha o arquivo `.env` fora do versionamento e use apenas `.env.example` no repositório.
> - Se houver exposição acidental de segredo em output/código, revogue e gere novas credenciais imediatamente.

[⬆ Voltar ao topo](#topo)

---

<a id="sec-executando-testes"></a>

## Executando testes

```bash
uv run pytest
```

Relatório de cobertura é impresso automaticamente. Os testes usam mocking e **não**
requerem credenciais reais.

[⬆ Voltar ao topo](#topo)

---

<a id="sec-criacao-em-lote"></a>

## Criação em lote de aplicações Azure AD

Para criar múltiplas aplicações Azure AD com credenciais automaticamente:

- **PowerShell (recomendado)**: `.\src/bulkCreate/Bulk-CreateApps.ps1 -InputPath config.json`
- **Python**: `python -m bulkCreate.bulk_create_apps config.json`

Autentique uma vez e execute múltiplas vezes com `-SkipLogin` (PowerShell) ou `--skip-login` (Python).
Use [examples/bulk_create_example.json](examples/bulk_create_example.json) como modelo e
veja [docs/bulk_create_apps.md](docs/bulk_create_apps.md) para documentação completa.

O utilitário em lote aplica o mesmo modelo de segurança do restante do projeto:

- exige `site_id` e `access_type` para toda aplicação
- usa `Sites.Selected` como `Role` para `app_only`
- usa `Sites.Selected` como `Scope` para `delegated`
- não adiciona autorizações de dados no nível do tenant

[⬆ Voltar ao topo](#topo)

---

<a id="sec-visao-geral-modulos"></a>

## Visão geral dos módulos

### `auth.py`
Obtém um token Bearer para a Microsoft Graph API usando o fluxo OAuth 2.0 de
**credenciais do cliente** via [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python).

Para o fluxo delegado, veja [docs/setup_delegated_auth.md](docs/setup_delegated_auth.md).

### `auth.py`
`GraphClient` é o cliente principal do Microsoft Graph. Ele gerencia a sessão
HTTP autenticada, expõe os helpers `get`, `post`, `patch`, `put_bytes` e
`get_raw`, e possui um `GraphAuthenticator` associado para descoberta do site.

### `drive.py`
Operações de biblioteca de documentos:

| Método | Descrição |
|---|---|
| `GraphDrive.list_drive_items(folder_path)` | Lista os filhos de uma pasta |
| `GraphDrive.download_file(item_id, local_path)` | Baixa um arquivo para disco |
| `GraphDrive.upload_file(local_path, remote_folder)` | Envia um arquivo local (≤ 4 MB) |
| `GraphDrive.read_file_content(item_id)` | Retorna o conteúdo textual do arquivo |
| `GraphDrive.write_file_content(item_id, content)` | Sobrescreve o conteúdo textual do arquivo |

### `lists.py`
Operações de listas do SharePoint:

| Método | Descrição |
|---|---|
| `GraphList.get_list_columns(names)` | Recupera colunas da lista, opcionalmente filtradas por nome |
| `GraphList.get_list_views()` | Lista as views da lista |
| `GraphList.get_list_view_columns(view_id)` | Lista as colunas visíveis em uma view |
| `GraphList.get_list_items(select, fields_only=False, include_title=False, include_item_id=False)` | Recupera itens da lista com seleção opcional de campos |
| `GraphList.create_list_item(fields)` | Cria um novo item |
| `GraphList.update_list_item(item_id, fields)` | Atualiza campos de um item existente |

### `auth.py`
`GraphAuthenticator` concentra a descoberta do site:

| Método | Descrição |
|---|---|
| `GraphAuthenticator.get_site_contents()` | Retorna metadados do site, drives e lists |
| `GraphAuthenticator.list_site_drives()` | Lista os drives do site |
| `GraphAuthenticator.list_site_lists()` | Lista as lists do site |

[⬆ Voltar ao topo](#topo)

---

<a id="sec-documentacao-adicional"></a>

## Documentação adicional

- [docs/getting_started.md](docs/getting_started.md) — visão geral, papéis administrativos e permissões
- [docs/setup_portal.md](docs/setup_portal.md) — configuração manual pelo portal
- [docs/setup_cli.md](docs/setup_cli.md) — configuração via Azure CLI e PowerShell
- [docs/setup_delegated_auth.md](docs/setup_delegated_auth.md) — fluxo delegado com login interativo
- [docs/bulk_create_apps.md](docs/bulk_create_apps.md) — criação em lote de aplicações

[⬆ Voltar ao topo](#topo)

---

<a id="sec-licenca"></a>

## Licença

Este projeto é licenciado sob a **GNU General Public License v3.0**.
Consulte [LICENSE](LICENSE) para o texto completo.

[⬆ Voltar ao topo](#topo)
