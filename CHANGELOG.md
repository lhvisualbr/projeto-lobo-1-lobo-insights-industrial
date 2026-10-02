# Changelog

> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

Registra somente o que foi realmente produzido em cada versão.

## V0.2.1 — Packaging & Reproducibility Fix

Microversão técnica de *release engineering* sobre a V0.2 auditada. **Nenhuma alteração funcional, analítica ou visual**: o XLSX atual é byte a byte igual ao XLSX auditado da V0.2 (SHA-256 `81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846`); dados, KPIs, fórmulas, rankings, Pareto, dashboard, gráficos e resultados não mudaram.

### Corrigido
- **Nomenclatura:** entregáveis principais com o prefixo obrigatório do Projeto 1 — ZIP `Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip` e XLSX `Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx`. Nomes históricos (`excel/lobo_insights_industrial_v0_1.xlsx`, `excel/lobo_insights_industrial_v0_2.xlsx`, scripts e documentos anteriores) foram preservados.
- **Manifesto de hashes:** `SHA256SUMS.txt` passa a listar só os arquivos **estáveis** e continua válido depois de executar os testes; os 4 arquivos regeneráveis (`docs/TEST_REPORT.md`, `images/dashboard_v0_2_pagina_1.png` a `_3.png`) ficam fora dele. Novo `RELEASE_SHA256SUMS.txt` = snapshot de todos os arquivos no instante do empacotamento.

### Produzido
- `scripts/lobo_release.py` (nomenclatura, manifestos, empacotamento determinístico), `scripts/construir_excel_v0_2_1.py` (gera o XLSX no nome novo, sobre o construtor da V0.2 inalterado), `scripts/executar_testes_v0_2_1.py` (executa a suíte da V0.1.1 e a bateria da V0.2 sem alteração, sobre o XLSX atual, e acrescenta os testes da V0.2.1).
- `docs/V0_2_1_NOTES.md`, `docs/V0_2_AUDITED_SHA256.txt` (30 arquivos da V0.2 auditada que não podem mudar), `docs/manual/observacoes_execucao_v0_2_1.md`; `README.md` e `docs/LIMITATIONS.md` atualizados; `docs/TEST_REPORT.md` regenerado para a V0.2.1.

### Validações executadas
- Ver `docs/TEST_REPORT.md` e `docs/V0_2_1_NOTES.md`: 201 testes (103 da V0.1.1 + 71 da V0.2 + 27 novos), incluindo testes de nomenclatura e controles positivos/negativos da estratégia de hashes.

### Não incluído nesta versão
- Qualquer lógica analítica nova, alteração do dashboard, previsão, IA, Power BI, n8n (V0.3 e posteriores).

## V0.2 — Analytics & Dashboard

Camada analítica e painel executivo em Excel sobre a baseline congelada V0.1.1. **Dados, regras, 6 KPIs e fórmulas existentes não foram alterados** (CSVs, scripts de regras e XLSX histórico byte a byte iguais). Sem previsão, sem IA.

### Produzido
- `excel/lobo_insights_industrial_v0_2.xlsx`: aba **DASHBOARD** (6 cartões de KPI, 9 gráficos, 6 cartões de concentração, fila de atenção de estoque, matriz criticidade × status, leituras por fórmula, impressão em 3 páginas); aba **Analises** (evolução mensal, categorias, centros, materiais, Pareto de custo e de consumo, concentração, fila de estoque, 18 conferências); 3 parâmetros novos em `Parametros` (A20:C24); `LEIA_ME` reescrita. Materiais, Estoque, Consumo, Qualidade e Indicadores idênticas à V0.1.1 (Indicadores muda só o texto de A1 e A3).
- `scripts/construir_excel_v0_2.py` (importa `construir()` da V0.1.1 e acrescenta as camadas) e `scripts/executar_testes_v0_2.py` (executa a suíte da V0.1.1 sem alteração e acrescenta 71 testes: T104–T174).
- Documentação: `docs/V0_2_NOTES.md`, `docs/V0_1_1_BASELINE_SHA256.txt`, `docs/TEST_REPORT_V0_1_1.md` (relatório da V0.1.1, preservado), `docs/manual/inspecao_visual_v0_2.md`, `docs/manual/observacoes_execucao_v0_2.md`; `README.md`, `docs/LIMITATIONS.md` e `docs/TEST_REPORT.md` atualizados; prévias em `images/`.

### Decisões
- Sem `SORT`/`FILTER`/`XLOOKUP` (ordenação por `RANK`+`COUNTIF`+`INDEX/MATCH`, com desempate determinístico — há empates reais nos dados).
- Filtros/segmentações avaliados e **não** implementados (justificativa em `docs/V0_2_NOTES.md`).

### Validações executadas
- Ver `docs/TEST_REPORT.md`: 174 testes (103 da V0.1.1 + 71 novos), incluindo cálculo independente de todas as tabelas da Analises, teste de mutação e reconstrução idêntica do XLSX.

### Não incluído nesta versão
- Previsão, tendência futura, IA/LLM, classificador, Power BI, n8n, API, banco de dados, aplicação web, GitHub; teste no Microsoft Excel real.

## V0.1.1 — Portable Baseline

Revisão técnica de estabilização e portabilidade. **Sem alteração funcional:** dados, KPIs, regras, fórmulas e estrutura do Excel permanecem idênticos aos da V0.1 (CSVs e XLSX byte a byte iguais).

### Problema identificado
- A geração do XLSX e parte da suíte de testes dependiam de um script fora do projeto (`recalc.py`, em pasta privada do ambiente original). Em outro ambiente, a suíte abortava. Não era erro de dados nem falha funcional.

### Correção realizada
- `scripts/construir_excel.py`: recálculo feito diretamente pelo LibreOffice em modo headless, localizado pela variável `LOBO_SOFFICE` ou pelo `PATH`; erro claro (`LibreOfficeIndisponivel`) quando ausente.
- `scripts/executar_testes.py`: testes que dependem do LibreOffice viram **SKIP** (nunca PASS) quando ele falta, em vez de abortar; código de saída 2 sinaliza suíte incompleta. Notas manuais do relatório movidas para dentro do pacote (`docs/manual/`). Corrigidos: f-string exclusiva do Python 3.12, import não usado e `open()` sem `with`.
- Novos testes T91 em diante (portabilidade, regressão contra a V0.1, qualidade do código). T01–T90 são os mesmos testes da V0.1.

### Arquivos novos
- `requirements.txt`, `docs/V0_1_1_NOTES.md`, `docs/V0_1_BASELINE_SHA256.txt`, `docs/TEST_REPORT_V0_1.md` (relatório original da V0.1, preservado), `docs/manual/inspecao_visual.md`, `docs/manual/observacoes_execucao.md`.

### Validações executadas
- Ver `docs/TEST_REPORT.md` (gerado nesta versão) e `docs/V0_1_1_NOTES.md`.

## V0.1 — Foundation

### Produzido
- Estrutura de pastas: `data/`, `excel/`, `docs/`, `scripts/` e pastas vazias para versões futuras (`prompts/`, `powerbi/`, `metrics/`, `n8n/`, `images/`).
- Três CSVs sintéticos: `materiais_ficticios.csv` (20 materiais), `estoque_ficticio.csv` (20 posições) e `consumo_ficticio.csv` (180 movimentações, 2026-06-01 a 2026-08-31).
- Excel `lobo_insights_industrial_v0_1.xlsx` com as abas LEIA_ME, Materiais, Estoque, Consumo, Qualidade, Indicadores e Parametros; tabelas estruturadas, fórmulas reais, validação de dados e formatação condicional.
- 36 verificações de qualidade, em Python e em fórmulas do Excel.
- 6 KPIs e resumos por mês, categoria, centro de trabalho e material, com conferências internas.
- Documentação: `README.md`, `DATA_DICTIONARY.md`, `LIMITATIONS.md`, `V0_1_NOTES.md`, `TEST_REPORT.md`.
- Scripts de geração, validação, construção do Excel e testes; `SHA256SUMS.txt`.

### Decisões e desvios documentados
- Interpretações adotadas para 10 pontos ambíguos do PDF (ver `docs/V0_1_NOTES.md`, seção 4).
- Pasta `scripts/` adicionada à estrutura do enunciado para garantir reprodutibilidade.

### Não incluído nesta versão
- Tabelas dinâmicas, dashboard, IA, classificador, avaliação, Power BI, n8n, publicação no GitHub e qualquer previsão.
