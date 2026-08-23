# Guia Arquitetural — Serviços (propostas de assessoria técnica)

**Howden Serviços · Propostas**
Aplicação em Blazor Server com persistência em DuckDB/Parquet, no mesmo padrão do
projeto **Licencas_HSA**.

---

## 1. Visão geral

O sistema substitui a planilha `Ferramenta para propostas de serviço.xlsm`, que hoje
concentra todo o ciclo comercial de serviço: lançar o **custo** (mão de obra e
despesas), calcular **margem e impostos** e emitir a **proposta** para o cliente.

As três guias da planilha viraram três telas encadeadas, e as macros VBA
(`MóduloProposta`, `MóduloCusto`, `MóduloWord`, `ClasseCusto`…) viraram código C#:

| Planilha | Aplicação |
|---|---|
| Guia **CUSTO** | `/servicos/custo` |
| Guia **PRICING** | `/servicos/pricing` |
| Guia **PROPOSTA** + macros de Word | `/servicos/proposta` |
| Guia **listas** (valores fixos) | `/servicos/parametros` (Tabela de Custos) |
| Cabeçalho/logo do documento | `/servicos/marca` (Identidade Visual) |
| Bloco de faturamento | `/servicos/faturamento` |
| — | `/servicos/propostas` (banco de propostas emitidas) |

### Stack

| Camada | Tecnologia |
|---|---|
| Runtime | .NET 8 (`net8.0`), ASP.NET Core |
| UI | Blazor Server (Razor Components, render interativo no servidor) |
| Dados | DuckDB em memória sobre arquivos Parquet (`DuckDB.NET.Data.Full` 1.1.3) |
| Importação | Leitura de `.xlsx`/`.xlsm` com `ClosedXML` 0.102.2 |
| Autenticação | Cookie (ASP.NET Core), credencial única da equipe |
| Estilo | CSS puro em `wwwroot/app.css` (o mesmo do projeto Licenças) |

---

## 2. Estrutura do projeto

```
servicos/
├── Program.cs                     # DI, autenticação, endpoints, seed
├── appsettings.json               # Pasta de dados, credencial, OpenBrowser
├── HowdenServicos.Poc.csproj      # net8.0 + DuckDB.NET + ClosedXML
├── Components/
│   ├── App.razor                  # Documento HTML raiz
│   ├── Routes.razor               # Router + AuthorizeRouteView (tudo exige login)
│   ├── RedirectToLogin.razor
│   ├── PaginaProposta.cs          # Base das 3 telas da proposta (rascunho + localStorage)
│   ├── Layout/                    # MainLayout, NavMenu, EmptyLayout
│   └── Pages/
│       ├── Home.razor             # "/" → /servicos/custo
│       ├── Login.razor
│       ├── Custo.razor            # mão de obra + despesas + importação da planilha
│       ├── PricingPage.razor      # riscos, provisões, margem alvo e impostos
│       ├── GerarProposta.razor    # cadastro, itens precificados, PDF/Word, gravação
│       ├── PropostasEnviadas.razor
│       ├── Parametros.razor       # tabela de custos padrão
│       ├── Marca.razor            # logo do documento
│       ├── Faturamento.razor      # dados de faturamento por BU
│       └── Error.razor
├── Data/
│   ├── ParquetStore.cs            # Núcleo da persistência (DuckDB sobre Parquet)
│   ├── Pricing.cs                 # ★ motor de cálculo (guias CUSTO + PRICING)
│   ├── Servicos.cs                # listas, rótulos PT/EN/ES e o documento HTML
│   ├── Rascunho.cs                # a proposta em edição (scoped no circuito)
│   ├── CustoImport.cs             # leitor da planilha antiga
│   ├── Repositorios.cs            # Proposta / Parametro / Faturamento / Branding
│   └── Seed.cs                    # valores padrão + DbInitializer
├── Models/                        # Proposta, ItemMO, ItemDespesa, PricingParams, …
└── wwwroot/
    ├── app.css                    # estilos (idênticos aos do Licenças + tabela de custo)
    └── app.js                     # imprimir, baixar arquivo e rascunho no localStorage
```

---

## 3. O motor de cálculo (`Data/Pricing.cs`)

É a tradução fiel da cadeia de preço da planilha. Partindo do custo:

```
custo (MO + despesas)
  → + risco de variação                        (PRICING E36/E38)
  → ÷ (1 − provisões − margem alvo)            (PRICING E52 — "montar preço a partir da margem")
  = VENDA LÍQUIDA (valor sem impostos)         (PRICING Q18)
  → ÷ (1 − PIS − COFINS − ISS)                 (PRICING Q20)
  = VALOR COM IMPOSTOS
```

- **Custo de um serviço** = `qtd. diárias × custo/hora × horas × (técnicos, se marcado)`
  — as linhas 8–17 da guia CUSTO, incluindo o `×$H$6` de quantidade de técnicos.
- **Custo de uma despesa** = `qtd × custo unitário × (técnicos, se marcado)` (linhas 22–31).
- **Provisões** (percentuais sobre a venda líquida): comissões, margem de negociação,
  PM + SACH, garantia do projeto e portal.
- **Margem alvo**: a planilha usava Goal Seek para achar o markup; aqui o preço sai
  direto da fórmula fechada — e o markup resultante é exibido para conferência
  (bate com o `E41` da planilha).
- **Impostos de serviço**: PIS, COFINS e ISS. Serviço não tem ICMS nem IPI — é o mesmo
  desvio que a planilha faz quando `Segmento = Service`.

### Distribuição do preço nos itens

A guia CUSTO tinha uma coluna de participação (`custo do item ÷ custo total`) que
multiplicava o preço de venda — é o que a coluna `L` fazia antes de a referência
quebrar (`#REF!`). Aqui isso está em `Pricing.Montar`:

- `valor do item = arredonda para cima (participação × valor com impostos)`
  (o `RoundUp` das classes `ClasseItemMO` / `ClasseItemDespesa`);
- `valor da diária = valor do item ÷ qtd. de diárias`;
- `valor da hora = valor da diária ÷ horas`;
- despesas: `valor unitário = arredonda para cima (valor do item ÷ qtd)`.

### Informações complementares (não incluso)

Mesma regra da `ClasseInfoComplementares` do VBA:

| Linha | Cálculo |
|---|---|
| DIARIA ADICIONAL | soma das diárias normais + unitários de hospedagem, carro, combustível e refeições |
| HORA EXTRA (dias úteis) | valor da hora do 1º turno × 1,5 |
| HORA EXTRA (sáb., dom. e feriados) | valor da hora do 1º turno × 2 |

### Conferência com a planilha

Cenário original do arquivo (24 diárias de 1 h a R$ 320 + R$ 585,36 de despesas,
risco 5%, comissões 1,65644%, MN 3%, PM+SACH 1%, garantia 2%, fiança R$ 19,03,
margem 36%, PIS 1,65%, COFINS 7,6%, ISS 7%):

| | Planilha | Aplicação |
|---|---|---|
| Custo total | 8.265,36 | 8.265,36 |
| Custo com riscos | 8.678,63 | 8.678,63 |
| Venda líquida (Q18) | 15.436,80 | 15.436,80 |
| Com PIS e COFINS (Q19) | 17.141,76 | 17.141,76 |
| Com PIS, COFINS e ISS (Q20) | 18.432,00 | 18.432,00 |
| Markup (E41) | 1,6403 | 1,6403 |
| Project Margin (D55) | 36,00% | 36,00% |

---

## 4. Camada de dados — Parquet + DuckDB

Igual ao projeto Licenças (`Data/ParquetStore.cs`, copiado sem alterações de lógica):

- **Escrita**: cada gravação vira um arquivo `.parquet` novo na subpasta da entidade,
  com as colunas de controle `_ts` (timestamp) e `_deleted` (exclusão lógica).
- **Leitura**: o DuckDB abre em memória, lê a pasta com `read_parquet(...)` e consolida
  com `row_number() OVER (PARTITION BY id ORDER BY _ts DESC)`.
- **Concorrência**: ninguém trava arquivo compartilhado — vários usuários gravam ao
  mesmo tempo na pasta de rede.

Entidades: `propostas`, `parametros` (tabela de custos, semeada), `faturamento`
(semeada com os dados da HSA-SP) e `branding` (logo).

A pasta vem de `Data:Folder` no `appsettings.json`; sem configuração, usa `data/`.

---

## 5. O rascunho da proposta

As três telas (Custo → Pricing → Proposta) editam **a mesma proposta em andamento**.
Ela vive em `Data/Rascunho.cs`, registrado como *scoped* — no Blazor Server isso
equivale ao circuito do usuário, então navegar pelo menu preserva tudo sem gravar
nada no banco.

Como um F5 (ou abrir `/servicos/pricing` direto na barra de endereços) começa um
circuito novo, `Components/PaginaProposta.cs` também grava o rascunho em JSON no
`localStorage` do navegador a cada alteração e o restaura na primeira renderização.
O botão **Nova proposta** limpa e recarrega os itens padrão da Tabela de Custos.

---

## 6. Autenticação

- Cookie com expiração deslizante de 7 dias.
- Credencial única da equipe em `appsettings.json` (`Auth:Usuario` / `Auth:Senha`);
  o login cria a identidade fixa "Equipe Howden", papel `admin`.
- `Routes.razor` usa `AuthorizeRouteView`; sem sessão, cai no `/login`.

> Nota: a senha fica em texto plano no `appsettings.json`, como na POC de Licenças.
> Para produção, mover para um segredo e considerar usuários individuais.

---

## 7. Rotas

| Rota | Página | O que faz |
|---|---|---|
| `/` | Home | Redireciona para `/servicos/custo` |
| `/login` | Login | Formulário que posta em `/auth/login` |
| `/servicos/custo` | Custo | Mão de obra e despesas; importa a planilha antiga |
| `/servicos/pricing` | Pricing | Riscos, provisões, margem alvo, impostos e composição do preço |
| `/servicos/proposta` | Gerar Proposta | Cadastro, itens precificados, PDF/Word e gravação |
| `/servicos/propostas` | Propostas Enviadas | Lista, busca, reabre e exporta CSV |
| `/servicos/parametros` | Tabela de Custos | Custo/hora e despesas padrão |
| `/servicos/marca` | Identidade Visual | Logo usado no documento |
| `/servicos/faturamento` | Dados de Faturamento | Razão social, endereço e banco por BU |
| `/servicos/propostas/export` | — | CSV das propostas (UTF-8 com BOM, separador `;`) |

---

## 8. O documento da proposta

`Servicos.DocBody` monta o HTML do documento com estilos inline (o Word exige isso),
compartilhado entre a impressão/PDF (`@media print` esconde a interface e mostra só
`.proposal-doc`) e o download `.doc`. Cores do modelo oficial: Arial, títulos em navy
`#141E32`, texto `#3C465A`, caixas e cabeçalhos de tabela em azul `#004785`.

Os rótulos são traduzidos em **português, inglês e espanhol** (`Servicos.Labels`),
como as macros `Macro_PT_br`, `Macro_Ingles` e `Macro_ESP` faziam.

---

## 9. O que ficou de fora (e por quê)

- **Envio de e-mail pelo Outlook** (`Módulo3_Automatização`): dependia do Outlook
  instalado na máquina. O caminho natural aqui é gerar o Word/PDF e anexar, ou
  configurar um SMTP — decidir com a equipe.
- **Guias ocultas de pricing de equipamento** (MP/MO de rotor, eixo, cubo, partes
  estáticas, back-to-back, ICMS/IPI por estado): são de proposta de produto, não de
  serviço. A tabela de ICMS por estado está preservada na planilha original caso o
  escopo cresça.
- **Controle de revisões** (guia `revisões`): hoje a revisão é um campo do cabeçalho;
  o histórico completo pode virar uma entidade nova quando for necessário.
