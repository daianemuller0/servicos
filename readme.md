# Howden Serviços · Propostas

Ferramenta web para propostas de **serviço / assessoria técnica** — a substituição
da planilha `Ferramenta para propostas de serviço.xlsm` (guias CUSTO, PRICING e
PROPOSTA + macros VBA).

Mesma stack e mesma identidade visual do projeto **Licencas_HSA**: .NET 8 com
Blazor Server, dados em Parquet numa pasta de rede consolidados pelo DuckDB,
login por cookie e CSS próprio.

## Rodar

```bash
dotnet run
```

Abre em <http://localhost:5081> (o navegador abre sozinho). Login padrão:
`howden` / `howden2026` — altere em `appsettings.json`.

Para servir a equipe a partir de uma máquina só:

```bash
HowdenServicos.Poc.exe --urls http://0.0.0.0:5081   # e "OpenBrowser": false
```

## O fluxo

Tudo em **uma tela** (Nova Proposta): dados do cliente e representante, margem,
riscos e impostos, custos de mão de obra e despesas — e embaixo a proposta já
precificada, com PDF/Word, e-mails prontos e gravação no banco.

Detalhes de arquitetura e das fórmulas: [ARQUITETURA.md](ARQUITETURA.md).
