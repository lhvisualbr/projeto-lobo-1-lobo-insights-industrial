# Lobo Insights Industrial

**Análise operacional, estoque e triagem inteligente — Projeto 1 de portfólio**

> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

**Status atual: V0.2.1 — Packaging & Reproducibility Fix**, microversão de empacotamento sobre a **V0.2 — Analytics & Dashboard** (auditada e aprovada). Nenhuma alteração funcional, analítica ou visual: o XLSX atual é byte a byte igual ao da V0.2. Corrige a nomenclatura dos entregáveis (prefixo `Projeto_Lobo_1_`) e torna o manifesto de hashes estável (ver `docs/V0_2_1_NOTES.md`).

**Entregáveis principais:** pacote `Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip`; Excel atual `Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx` (em `excel/`).

**V0.2 — Analytics & Dashboard.** Camada analítica e aba `DASHBOARD` (KPIs em cartões, 9 gráficos, dois Paretos, situação de estoque e criticidade) construídas **sobre a baseline congelada V0.1.1**, sem alterar dados, regras, KPIs ou fórmulas existentes. O XLSX histórico da V0.1.1 foi preservado. IA, classificador, Power BI, automação e **qualquer previsão ainda não existem** e não são descritos como prontos.

## Objetivo

Construir, por versões pequenas e verificáveis, um estudo de caso industrial fictício que responda:

> **"O que aconteceu na operação e o que precisa de atenção?"**

O projeto descreve o **presente e o histórico** (consumo, custo estimado, estoque atual). Ele **não faz previsão de demanda nem cálculo preditivo de reposição** — isso pertence ao Projeto 2 (Lobo Forecast AI Industrial).

## Contexto fictício

Uma operação industrial inventada consome materiais de uso diário (abrasivos, soldagem, corte, elétrica, EPI e ferramentas). Os registros estão em três bases: catálogo de materiais, posição de estoque e movimentações de consumo de três meses consecutivos (junho a agosto de 2026). Nenhuma empresa, pessoa, fornecedor, valor ou processo real foi usado; fornecedores são "Fornecedor Alfa/Beta/Gama/Delta/Sigma".

## Tecnologias realmente usadas (V0.1 a V0.2)

| Tecnologia | Uso |
|---|---|
| Python 3 (`pandas`, `openpyxl`, `lxml`) | Gerar os CSVs, validar a qualidade, construir o Excel (incluindo os gráficos nativos do Excel) e rodar os testes |
| Excel (.xlsx) | Tabelas estruturadas, fórmulas (SOMASES, CONT.SES, SE, SEERRO, ÍNDICE/CORRESP, SOMARPRODUTO), validação de dados, formatação condicional |
| LibreOffice (apenas verificação) | Recalcular e renderizar o XLSX durante os testes |
| `pdftoppm` / `pdfinfo` (poppler, opcional) | Testes de renderização do DASHBOARD (V0.2); sem eles os testes ficam SKIP |

Não usados: Power BI, n8n, IA/LLM, GitHub, Python de previsão, SQL. Nenhuma biblioteca Python nova foi adicionada na V0.2.

## Estrutura

```
lobo-insights-industrial/
├── README.md
├── CHANGELOG.md
├── data/        materiais_ficticios.csv, estoque_ficticio.csv, consumo_ficticio.csv        (inalterados desde a V0.1)
├── excel/       Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx   (XLSX ATUAL; igual ao auditado da V0.2)
│                lobo_insights_industrial_v0_2.xlsx                    (histórico V0.2, preservado)
│                lobo_insights_industrial_v0_1.xlsx                    (histórico V0.1/V0.1.1, preservado)
├── docs/        DATA_DICTIONARY.md, LIMITATIONS.md, TEST_REPORT.md (regenerável), V0_2_1_NOTES.md, V0_2_NOTES.md,
│                V0_1_NOTES.md, V0_1_1_NOTES.md, TEST_REPORT_V0_1.md, TEST_REPORT_V0_1_1.md (históricos),
│                V0_1_BASELINE_SHA256.txt, V0_1_1_BASELINE_SHA256.txt, V0_2_AUDITED_SHA256.txt, manual/
├── scripts/     gerar_dados.py, validar_dados.py, lobo_common.py, construir_excel.py, executar_testes.py   (V0.1.1)
│                construir_excel_v0_2.py, executar_testes_v0_2.py                                            (V0.2)
│                lobo_release.py, construir_excel_v0_2_1.py, executar_testes_v0_2_1.py                      (V0.2.1)
├── images/      dashboard_v0_2_pagina_1.png … _3.png  (prévias regeneráveis)
├── requirements.txt
├── prompts/  powerbi/  metrics/  n8n/    (vazias: versões futuras)
├── SHA256SUMS.txt            manifesto operacional (arquivos estáveis; passa depois de executar os testes)
└── RELEASE_SHA256SUMS.txt    snapshot da release (todos os arquivos; confere só no pacote recém-extraído)
```

A pasta `scripts/` não estava na estrutura do enunciado original; foi adicionada na V0.1 para tornar o pacote **reproduzível** (ver `docs/V0_1_NOTES.md`).

## Como abrir

- **Excel:** abra `excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx` (idêntico ao histórico `excel/lobo_insights_industrial_v0_2.xlsx` da V0.2). A primeira aba é o **`DASHBOARD`** (leitura executiva). Abas: `DASHBOARD`, `LEIA_ME`, `Indicadores` (6 KPIs oficiais), `Analises` (camada de cálculo: rankings, Pareto, concentração, fila de estoque, conferências), `Materiais`, `Estoque`, `Consumo`, `Qualidade`, `Parametros`. Todos os números do painel são fórmulas: `tbConsumo/tbEstoque/tbMateriais → Indicadores → Analises → DASHBOARD`.
- **Prévias:** `images/dashboard_v0_2_pagina_1.png` a `_3.png` (regeneráveis) (renderizadas pelo LibreOffice; no Excel a aparência pode diferir — ver limitações).
- **CSVs:** separador `;`, UTF-8 com BOM, decimal com **ponto** (ex.: `9.80`). No Excel em português, importe por *Dados > De Texto/CSV* e escolha a localidade que use ponto decimal, ou simplesmente use o XLSX pronto.

## Integridade dos arquivos (SHA-256)

Dois manifestos com papéis diferentes (detalhes em `docs/V0_2_1_NOTES.md`):

- `SHA256SUMS.txt` — **operacional**: só arquivos estáveis (dados, scripts, XLSX, documentação-fonte). `sha256sum -c SHA256SUMS.txt` continua passando **depois** de executar os testes. Os 4 arquivos regeneráveis (`docs/TEST_REPORT.md`, `images/dashboard_v0_2_pagina_1.png` a `_3.png`) ficam fora dele de propósito; qualquer outro arquivo é controlado.
- `RELEASE_SHA256SUMS.txt` — **snapshot** de todos os arquivos no instante do empacotamento; confere integralmente só no pacote recém-extraído.

Utilitário: `python scripts/lobo_release.py verificar | manifestos | empacotar`.

## Como validar

```bash
pip install -r requirements.txt             # pandas, openpyxl, lxml (versões testadas)
python scripts/validar_dados.py                      # 36 verificações de qualidade nos CSVs
python scripts/executar_testes.py --verificar        # suíte da V0.1.1 (103 testes; não sobrescreve o relatório)
python scripts/executar_testes_v0_2.py --verificar   # V0.2 (histórico): suíte da V0.1.1 sem alteração + 71 testes (174 no total)
python scripts/executar_testes_v0_2_1.py --verificar # V0.2.1 (atual): as duas anteriores sem alteração, sobre o XLSX atual, + 27 testes (201 no total)
```

Os scripts funcionam a partir de qualquer diretório (os caminhos são resolvidos a partir da pasta do projeto). O **LibreOffice** é uma ferramenta externa necessária apenas para reconstruir o XLSX e para os testes que recalculam fórmulas: deixe o comando `soffice` no `PATH` ou defina `LOBO_SOFFICE` com o caminho do executável. Sem ele, esses testes aparecem como **SKIP** (nunca PASS) e a suíte termina com código de saída 2; código 0 = tudo executado e aprovado, 1 = há FAIL. Ambiente verificado: Linux x86_64, Python 3.12.3 (ver `docs/V0_1_1_NOTES.md`).

No Excel: aba `Qualidade` (todos PASS) e aba `Indicadores`, seção 6 (todas as conferências OK). O resultado detalhado (regenerável) está em `docs/TEST_REPORT.md`; o da V0.1.1, preservado, em `docs/TEST_REPORT_V0_1_1.md`.

Para recriar tudo do zero: `python scripts/gerar_dados.py` (mesma semente = mesmos CSVs), `python scripts/construir_excel.py` (XLSX histórico V0.1.1) e `python scripts/construir_excel_v0_2.py` (XLSX V0.2); ambos requerem LibreOffice para gravar os valores calculados dentro do arquivo. `executar_testes_v0_2_1.py` (sem `--verificar`) regrava os 4 arquivos regeneráveis: `docs/TEST_REPORT.md` e as prévias em `images/`.

## Indicadores oficiais — V0.1 a V0.2 (valores de referência de regressão)

Calculados a partir dos CSVs e reconciliados com o Excel V0.1.1, com a aba `Indicadores` da V0.2 e com os cartões do `DASHBOARD` (ver `docs/TEST_REPORT.md`). Valores monetários são fictícios.

| Indicador | Valor |
|---|---|
| Unidades consumidas | 1.647 |
| Custo estimado consumido | 24.606,20 |
| Movimentações | 180 |
| Materiais distintos movimentados | 20 |
| Itens abaixo ou iguais ao mínimo | 5 |
| Itens críticos abaixo ou iguais ao mínimo | 2 |

Status de estoque (regra simples de estado atual, não é previsão; inalterada): **REPOR** se `Estoque_Atual <= Estoque_Minimo`; **ATENÇÃO** se acima do mínimo, mas até 25% acima dele; **OK** caso contrário.

## Roadmap

| Versão | Foco |
|---|---|
| **V0.1** | Foundation: dados sintéticos, validações, Excel base |
| **V0.1.1** | Portable Baseline: portabilidade e reprodutibilidade, sem mudança funcional |
| **V0.2** | Analytics & Dashboard: camada analítica, Pareto, estoque e criticidade, `DASHBOARD` — **atual** |
| V0.3 | IA executiva (insights a partir de agregados) |
| V0.4 | Classificador de solicitações (60 mensagens, Prompt V1/V2) |
| V0.5 | Avaliação das métricas e análise de erros |
| V0.6 | Power BI (modelo, DAX, 2 páginas) |
| V0.7 | Automação n8n com revisão humana |
| V1.0 | Portfólio: README final, imagens, release |

## Aviso

Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio. Nenhum uso operacional real. Limitações em `docs/LIMITATIONS.md`.
