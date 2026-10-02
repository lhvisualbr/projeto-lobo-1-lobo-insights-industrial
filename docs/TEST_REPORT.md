# TEST_REPORT — Projeto Lobo 1 — Lobo Insights Industrial V0.2.1 (Packaging & Reproducibility Fix)
> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

Relatório **gerado por `scripts/executar_testes_v0_2_1.py`**; cada PASS/FAIL vem de uma execução real e teste que não pôde rodar aparece como **SKIP** (nunca PASS). **Este arquivo é regenerável**: registra o ambiente de execução (versões e plataforma) e por isso não faz parte do manifesto operacional `SHA256SUMS.txt` (ver `docs/V0_2_1_NOTES.md`). A V0.2.1 é uma microversão de empacotamento: **nenhuma lógica analítica, dado, fórmula, gráfico ou layout mudou** — o XLSX atual é byte a byte igual ao XLSX auditado da V0.2. Históricos preservados: `docs/TEST_REPORT_V0_1.md`, `docs/TEST_REPORT_V0_1_1.md` (o relatório da V0.2 auditada está descrito em `docs/V0_2_NOTES.md`). Para reproduzir: `python scripts/executar_testes_v0_2_1.py --verificar`.

## 1. Resumo
- Testes automatizados: **201** — PASS **201**, FAIL **0**, SKIP **0**.
- Suíte da V0.1.1 (T01–T103, sem alteração): Testes automatizados: 103 | PASS 103 | FAIL 0 | SKIP 0 | Verificações de qualidade (Python): 36 | FAIL 0 | Detecção por mutação: 36 | FAIL 0 | SKIP 0
- Bateria da V0.2 (T104–T174, sem alteração, **executada sobre `excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx`**): **71** — PASS **71**, FAIL **0**, SKIP **0**.
- Testes novos da V0.2.1 (T175–T201): **27** — PASS **27**, FAIL **0**, SKIP **0**.
- Verificações de qualidade dos CSVs (Python, Q01–Q36): **36** — PASS **36**, FAIL **0**.
- Ambiente: Python 3.12.3, pandas 3.0.2, openpyxl 3.1.5, LibreOffice 24.2.7.2 420(Build:2), Linux x86_64.
- SHA-256 do XLSX atual (`excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx`): `81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846`
- SHA-256 do XLSX auditado da V0.2 (`excel/lobo_insights_industrial_v0_2.xlsx`, nome histórico preservado): `81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846` — **idêntico**.
- SHA-256 dos CSVs (inalterados): `materiais_ficticios.csv` = `eca98597dc740956b4db95885ec54693def431a56501ea6d055fa54c1654231b`; `estoque_ficticio.csv` = `9f90cd52f208f9f863b2ababf7fed4735bbe893ef137efdc79f7506b3359a8d0`; `consumo_ficticio.csv` = `b5b8249731e5fcbce34b4345b4d1e576c568b33de3eba4dfca91f8bd0d94233c`

## 2. TESTADO × NÃO TESTADO
**TESTADO (executado):** toda a suíte da V0.1.1; toda a bateria da V0.2 (KPIs, 36 verificações de qualidade, 2.417 fórmulas sem erro, sem vínculos externos, 9 gráficos, cálculo independente de todas as tabelas da aba Analises, mutação, reconstrução, renderização) **agora sobre o XLSX com o nome da V0.2.1**; nomenclatura obrigatória (XLSX, ZIP, rejeição de nomes ambíguos, nomes históricos preservados); igualdade byte a byte do XLSX atual com o auditado; 30 arquivos da V0.2 auditada inalterados; manifesto operacional (formato, ordem, cobertura, reprodutibilidade, `sha256sum -c`); snapshot da release; **controles positivos e negativos da estratégia de hashes** em cópia descartável; documentação.

**NÃO TESTADO (declarado, não assumido):**
- **Microsoft Excel real** (Windows/Mac/Online) e **Excel em pt-BR**: mesmas limitações da V0.2 (`docs/LIMITATIONS.md`, itens 17–25). Como o XLSX é byte a byte o mesmo da V0.2, nada novo foi verificado nem quebrado neste ponto.
- **Variação real do ambiente do auditor:** a falha de manifesto relatada **não foi reproduzida** aqui (nesta máquina as regenerações saem byte a byte idênticas); a estratégia foi validada por simulação — os 4 regeneráveis foram alterados numa cópia (ver seção 3) — e pela execução completa dos testes seguida de `sha256sum -c`, registradas em `docs/manual/observacoes_execucao_v0_2_1.md`.
- **O ZIP final** só pode ser conferido após o empacotamento (um arquivo não contém a prova do próprio ZIP): essa conferência (nome, conteúdo, `RELEASE_SHA256SUMS.txt`, execução a partir de outro diretório) é feita fora desta suíte e registrada nas observações.
- Instalação limpa por `pip install -r requirements.txt` (sem rede neste ambiente); Windows/macOS; outras versões de Python/LibreOffice.
- Inspeção visual humana: nenhuma alteração visual foi feita; vale a inspeção da V0.2 (seção 8), parcial.

## 3. Manifestos SHA-256 e arquivos regeneráveis
- `SHA256SUMS.txt` — manifesto **operacional**: todos os arquivos estáveis (dados, scripts, XLSX, documentação-fonte, manifestos históricos), sem os 4 regeneráveis. Continua válido depois de executar os testes.
- `RELEASE_SHA256SUMS.txt` — **snapshot** de todos os arquivos (inclusive `SHA256SUMS.txt` e os regeneráveis) no instante do empacotamento; só confere integralmente no pacote recém-extraído.
- Regeneráveis: `docs/TEST_REPORT.md`, `images/dashboard_v0_2_pagina_1.png`, `images/dashboard_v0_2_pagina_2.png`, `images/dashboard_v0_2_pagina_3.png`.

Controles da estratégia (cópia descartável do projeto):

| Situação simulada | Verificação | Resultado |
|---|---|---|
| 4 regeneráveis alterados | SHA256SUMS passa | PASS |
| 4 regeneráveis alterados | snapshot acusa exatamente os 4 | PASS |
| 1 imagem regenerável apagada | SHA256SUMS passa | PASS |
| alteração real em data/consumo_ficticio.csv | SHA256SUMS reprova só esse arquivo | PASS |
| alteração real em scripts/construir_excel_v0_2.py | SHA256SUMS reprova só esse arquivo | PASS |
| alteração real em excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx | SHA256SUMS reprova só esse arquivo | PASS |
| alteração real em docs/V0_2_NOTES.md | SHA256SUMS reprova só esse arquivo | PASS |
| arquivo estável apagado | SHA256SUMS acusa ausente | PASS |
| arquivo novo não registrado | teste de cobertura o detecta | PASS |

## 4. Verificações de qualidade dos dados (Python, sobre os CSVs)
| Teste | Regra | Resultado | Quantidade de erros | Status |
|---|---|---|---|---|
| Q01 · IDs duplicados – Materiais | Material_ID não se repete (conta linhas envolvidas em duplicidade) | Nenhum erro encontrado | 0 | PASS |
| Q02 · IDs duplicados – Estoque | Material_ID não se repete na posição de estoque | Nenhum erro encontrado | 0 | PASS |
| Q03 · IDs duplicados – Consumo | Movimento_ID não se repete | Nenhum erro encontrado | 0 | PASS |
| Q04 · Nulos indevidos – Materiais | Nenhuma célula vazia na tabela | Nenhum erro encontrado | 0 | PASS |
| Q05 · Nulos indevidos – Estoque | Nenhuma célula vazia na tabela | Nenhum erro encontrado | 0 | PASS |
| Q06 · Nulos indevidos – Consumo | Nenhuma célula vazia na tabela | Nenhum erro encontrado | 0 | PASS |
| Q07 · Quantidade inválida – Consumo | Quantidade numérica, inteira e > 0 | Nenhum erro encontrado | 0 | PASS |
| Q08 · Custo <= 0 – Materiais | Custo_Unitario numérico e > 0 | Nenhum erro encontrado | 0 | PASS |
| Q09 · Custo <= 0 – Estoque | Custo_Unitario numérico e > 0 | Nenhum erro encontrado | 0 | PASS |
| Q10 · Custo <= 0 – Consumo | Custo_Unitario numérico e > 0 | Nenhum erro encontrado | 0 | PASS |
| Q11 · Material inexistente – Estoque | Material_ID do estoque existe no catálogo | Nenhum erro encontrado | 0 | PASS |
| Q12 · Material inexistente – Consumo | Material_ID do consumo existe no catálogo | Nenhum erro encontrado | 0 | PASS |
| Q13 · Catálogo sem estoque | Todo material do catálogo tem posição de estoque | Nenhum erro encontrado | 0 | PASS |
| Q14 · Nome divergente – Estoque | Material do estoque = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q15 · Nome divergente – Consumo | Material do consumo = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q16 · Categoria divergente – Estoque | Categoria do estoque = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q17 · Categoria divergente – Consumo | Categoria do consumo = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q18 · Unidade divergente – Estoque | Unidade do estoque = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q19 · Unidade divergente – Consumo | Unidade do consumo = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q20 · Custo divergente – Estoque | Custo_Unitario do estoque = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q21 · Custo divergente – Consumo | Custo_Unitario do consumo = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q22 · Criticidade divergente – Estoque | Criticidade do estoque = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q23 · Fornecedor divergente – Estoque | Fornecedor do estoque = catálogo | Nenhum erro encontrado | 0 | PASS |
| Q24 · Datas inválidas – Consumo | Data válida e dentro do período definido na aba Parametros | Nenhum erro encontrado | 0 | PASS |
| Q25 · Cobertura dos 3 meses – Consumo | Cada mês do período tem >= 1 movimentação (conta meses sem registro) | Nenhum erro encontrado | 0 | PASS |
| Q26 · Estoque negativo | Estoque_Atual numérico e >= 0 | Nenhum erro encontrado | 0 | PASS |
| Q27 · Estoque mínimo negativo | Estoque_Minimo numérico e >= 0 | Nenhum erro encontrado | 0 | PASS |
| Q28 · Lead time inválido | Lead_Time_Dias numérico, inteiro e > 0 | Nenhum erro encontrado | 0 | PASS |
| Q29 · Categoria inválida – Materiais | Categoria pertence à lista permitida | Nenhum erro encontrado | 0 | PASS |
| Q30 · Categoria inválida – Estoque | Categoria pertence à lista permitida | Nenhum erro encontrado | 0 | PASS |
| Q31 · Categoria inválida – Consumo | Categoria pertence à lista permitida | Nenhum erro encontrado | 0 | PASS |
| Q32 · Unidade inválida – Materiais | Unidade em UN, CX, PCT ou PAR | Nenhum erro encontrado | 0 | PASS |
| Q33 · Criticidade inválida – Materiais | Criticidade em Baixa, Média ou Alta | Nenhum erro encontrado | 0 | PASS |
| Q34 · Centro inválido – Consumo | Centro_Trabalho na lista permitida | Nenhum erro encontrado | 0 | PASS |
| Q35 · Tipo de movimentação inválido | Tipo_Movimentacao em Consumo ou Ajuste | Nenhum erro encontrado | 0 | PASS |
| Q36 · Fornecedor fora do padrão fictício | Fornecedor começa com 'Fornecedor ' (nome fictício) | Nenhum erro encontrado | 0 | PASS |

## 5. Testes automatizados (T01–T103 V0.1.1; T104–T174 V0.2; T175 em diante V0.2.1)
| ID | Área | Cenário | Esperado | Obtido | PASS/FAIL |
|---|---|---|---|---|---|
| T01 | Linhas | materiais_ficticios.csv – nº de linhas de dados | 20 | 20 | PASS |
| T02 | Linhas | estoque_ficticio.csv – nº de linhas de dados | 20 | 20 | PASS |
| T03 | Linhas | consumo_ficticio.csv – nº de movimentações | 180 | 180 | PASS |
| T04 | Linhas | Excel: linhas de tbMateriais / tbEstoque / tbConsumo = CSVs | 20/20/180 | 20/20/180 | PASS |
| T05 | Unicidade de IDs | Material_ID único no catálogo (distintos = linhas) | 20 | 20 | PASS |
| T06 | Unicidade de IDs | Material_ID único no estoque | 20 | 20 | PASS |
| T07 | Unicidade de IDs | Movimento_ID único no consumo | 180 | 180 | PASS |
| T08 | Unicidade de IDs | Material_ID = MAT001..MAT020 | True | True | PASS |
| T09 | Unicidade de IDs | Movimento_ID = MOV0001..MOV0180 (sem lacunas) | True | True | PASS |
| T10 | Integridade referencial | consumo -> catálogo: Material_ID órfãos | 0 | 0 | PASS |
| T11 | Integridade referencial | estoque -> catálogo: Material_ID órfãos | 0 | 0 | PASS |
| T12 | Integridade referencial | catálogo -> estoque: materiais sem posição | 0 | 0 | PASS |
| T13 | Integridade referencial | nome/categoria/unidade/custo/criticidade/fornecedor divergentes (Q14–Q23) | 0 | 0 | PASS |
| T14 | Custos | custos > 0 nas três tabelas (Q08–Q10) | 0 | 0 | PASS |
| T15 | Quantidades | quantidade inteira > 0 em todas as movimentações (Q07); mín/máx observados | 0 erros | 0 erros (mín 1, máx 60) | PASS |
| T16 | Datas | todas as datas válidas e entre 2026-06-01 e 2026-08-31 (Q24) | 0 erros | 0 erros (2026-06-01 a 2026-08-31) | PASS |
| T17 | Datas | três meses consecutivos com movimentação (Q25) | ['2026-06', '2026-07', '2026-08'] | ['2026-06', '2026-07', '2026-08'] | PASS |
| T18 | Categorias | categorias válidas nas três tabelas (Q29–Q31) | 0 | 0 | PASS |
| T19 | Categorias | as 6 categorias do PDF estão representadas no catálogo | ['Abrasivos', 'Corte', 'EPI', 'Elétrica', 'Ferramentas', 'Soldagem'] | ['Abrasivos', 'Corte', 'EPI', 'Elétrica', 'Ferramentas', 'Soldagem'] | PASS |
| T20 | Centros | centros de trabalho na lista permitida (Q34) | 0 | 0 | PASS |
| T21 | Centros | os 5 centros permitidos aparecem no consumo | ['ELETRICA', 'FABRICACAO', 'MANUTENCAO', 'MECANICA', 'SOLDAGEM'] | ['ELETRICA', 'FABRICACAO', 'MANUTENCAO', 'MECANICA', 'SOLDAGEM'] | PASS |
| T22 | Estoque | estoque >= 0, mínimo >= 0, lead time inteiro > 0 (Q26–Q28) | 0 | 0 | PASS |
| T23 | Estoque | há itens com Atual <= Mínimo (REPOR) | >= 1 | 5 | PASS |
| T24 | Estoque | PDF: pelo menos 5 itens em ATENÇÃO/REPOR | >= 5 | 9 | PASS |
| T25 | Estoque | há ao menos um item de estoque exatamente igual ao mínimo | >= 1 | 1 | PASS |
| T26 | Estoque | criticidades Baixa, Média e Alta presentes no estoque | 3 níveis | {'Média': 9, 'Baixa': 6, 'Alta': 5} | PASS |
| T27 | Estoque | pelo menos um item Alta abaixo ou igual ao mínimo (cenário para o KPI 6) | >= 1 | 2 | PASS |
| T28 | Distribuição | consumo NÃO é uniforme entre materiais (máx/mín de movimentações >= 3) | >= 3 | 22.0 | PASS |
| T29 | Distribuição | consumo NÃO é uniforme entre meses (movimentações por mês distintas) | >= 2 valores distintos | {'2026-06': 65, '2026-07': 75, '2026-08': 40} | PASS |
| T30 | Formato CSV | materiais_ficticios.csv: cabeçalho exato | ['Material_ID', 'Material', 'Categoria', 'Unidade', 'Custo_Unitario', 'Criticidade', 'Fornecedor'] | ['Material_ID', 'Material', 'Categoria', 'Unidade', 'Custo_Unitario', 'Criticidade', 'Fornecedor'] | PASS |
| T31 | Formato CSV | materiais_ficticios.csv: separador ';' e nº de campos constante | {7} | {7} | PASS |
| T32 | Formato CSV | materiais_ficticios.csv: UTF-8 com BOM, finais de linha LF | BOM=True, CRLF=False | BOM=True, CRLF=False | PASS |
| T33 | Formato CSV | estoque_ficticio.csv: cabeçalho exato | ['Material_ID', 'Material', 'Categoria', 'Unidade', 'Estoque_Atual', 'Estoque_Minimo', 'Custo_Unitario', 'Lead_Time_Dias', 'Fornecedor', 'Criticidade'] | ['Material_ID', 'Material', 'Categoria', 'Unidade', 'Estoque_Atual', 'Estoque_Minimo', 'Custo_Unitario', 'Lead_Time_Dias', 'Fornecedor', 'Criticidade'] | PASS |
| T34 | Formato CSV | estoque_ficticio.csv: separador ';' e nº de campos constante | {10} | {10} | PASS |
| T35 | Formato CSV | estoque_ficticio.csv: UTF-8 com BOM, finais de linha LF | BOM=True, CRLF=False | BOM=True, CRLF=False | PASS |
| T36 | Formato CSV | consumo_ficticio.csv: cabeçalho exato | ['Movimento_ID', 'Data', 'Material_ID', 'Material', 'Categoria', 'Quantidade', 'Unidade', 'Custo_Unitario', 'Centro_Trabalho', 'Tipo_Movimentacao'] | ['Movimento_ID', 'Data', 'Material_ID', 'Material', 'Categoria', 'Quantidade', 'Unidade', 'Custo_Unitario', 'Centro_Trabalho', 'Tipo_Movimentacao'] | PASS |
| T37 | Formato CSV | consumo_ficticio.csv: separador ';' e nº de campos constante | {10} | {10} | PASS |
| T38 | Formato CSV | consumo_ficticio.csv: UTF-8 com BOM, finais de linha LF | BOM=True, CRLF=False | BOM=True, CRLF=False | PASS |
| T39 | Formato CSV | Custo_Unitario com ponto decimal e 2 casas nos 3 CSVs | 0 | 0 | PASS |
| T40 | Reprodutibilidade | gerar_dados.py (mesma semente) reproduz os 3 CSVs byte a byte | True | True | PASS |
| T41 | Excel – estrutura | abas na ordem | ['LEIA_ME', 'Materiais', 'Estoque', 'Consumo', 'Qualidade', 'Indicadores', 'Parametros'] | ['LEIA_ME', 'Materiais', 'Estoque', 'Consumo', 'Qualidade', 'Indicadores', 'Parametros'] | PASS |
| T42 | Excel – estrutura | tabelas estruturadas e intervalos | {'tbMateriais': 'A1:G21', 'tbEstoque': 'A1:L21', 'tbConsumo': 'A1:L181', 'tbQualidade': 'A9:E45'} | {'tbMateriais': 'A1:G21', 'tbEstoque': 'A1:L21', 'tbConsumo': 'A1:L181', 'tbQualidade': 'A9:E45'} | PASS |
| T43 | Excel – estrutura | nomes definidos (listas e parâmetros) | ['lstCategorias', 'lstCentros', 'lstCriticidade', 'lstMeses', 'lstTipoMov', 'lstUnidades', 'pFimPeriodo', 'pInicioPeriodo', 'pMargemAtencao'] | ['lstCategorias', 'lstCentros', 'lstCriticidade', 'lstMeses', 'lstTipoMov', 'lstUnidades', 'pFimPeriodo', 'pInicioPeriodo', 'pMargemAtencao'] | PASS |
| T44 | Excel – estrutura | sem vínculos externos (externalLinks) | 0 | 0 | PASS |
| T45 | Excel – estrutura | recálculo total ao abrir (fullCalcOnLoad) | True | True | PASS |
| T46 | Excel – estrutura | validações de dados por aba (Materiais/Estoque/Consumo) | 4/6/8 | 4/6/8 | PASS |
| T47 | Excel – estrutura | intervalos com formatação condicional (Materiais/Estoque/Qualidade/Indicadores) | 1/3/1/3 | 1/3/1/3 | PASS |
| T48 | Excel – fórmulas | fórmulas reais no arquivo (>= 800) | >= 800 | 857 | PASS |
| T49 | Excel – fórmulas | células de fórmula sem valor em cache | 0 | 0 | PASS |
| T50 | Excel – fórmulas | células com erro (#VALUE!, #NAME?, #REF!…) nos valores calculados | 0 | 0 | PASS |
| T51 | Excel – fórmulas | funções do PDF presentes (SOMASES, CONT.SES, SE, SEERRO, ÍNDICE/CORRESP…) | todas > 0 | {'SUMIFS': 70, 'COUNTIFS': 46, 'COUNTIF': 40, 'IF': 185, 'IFERROR': 100, 'INDEX': 100, 'MATCH': 100, 'SUMPRODUCT': 28, 'SUM': 16, 'ROUND': 180, 'COUNTBLANK': 3} | PASS |
| T52 | Excel – fórmulas | aba Indicadores: números digitados à mão (constantes numéricas) | 0 | 0 | PASS |
| T53 | Excel – fórmulas | Custo_Movimentado = Quantidade x Custo_Unitario nas 180 linhas | 180 | 180 | PASS |
| T54 | Excel – fórmulas | Mes = aaaa-mm da Data nas 180 linhas | 180 | 180 | PASS |
| T55 | Excel – fórmulas | Status_Estoque (Excel) = regra em Python nas 20 linhas | 20 | 20 | PASS |
| T56 | Excel – fórmulas | Critico_Em_Risco (Excel) = regra em Python nas 20 linhas | 20 | 20 | PASS |
| T57 | Excel – fórmulas | aba Qualidade: vetor de erros do Excel (Q01–Q36) = vetor do Python | {'Q01': 0, 'Q02': 0, 'Q03': 0, 'Q04': 0, 'Q05': 0, 'Q06': 0, 'Q07': 0, 'Q08': 0, 'Q09': 0, 'Q10': 0, 'Q11': 0, 'Q12': 0, 'Q13': 0, 'Q14': 0, 'Q15': 0, 'Q16': 0, 'Q17': 0, 'Q18': 0, 'Q19': 0, 'Q20': 0, 'Q21': 0, 'Q22': 0, 'Q23': 0, 'Q24': 0, 'Q25': 0, 'Q26': 0, 'Q27': 0, 'Q28': 0, 'Q29': 0, 'Q30': 0, 'Q31': 0, 'Q32': 0, 'Q33': 0, 'Q34': 0, 'Q35': 0, 'Q36': 0} | {'Q01': 0, 'Q02': 0, 'Q03': 0, 'Q04': 0, 'Q05': 0, 'Q06': 0, 'Q07': 0, 'Q08': 0, 'Q09': 0, 'Q10': 0, 'Q11': 0, 'Q12': 0, 'Q13': 0, 'Q14': 0, 'Q15': 0, 'Q16': 0, 'Q17': 0, 'Q18': 0, 'Q19': 0, 'Q20': 0, 'Q21': 0, 'Q22': 0, 'Q23': 0, 'Q24': 0, 'Q25': 0, 'Q26': 0, 'Q27': 0, 'Q28': 0, 'Q29': 0, 'Q30': 0, 'Q31': 0, 'Q32': 0, 'Q33': 0, 'Q34': 0, 'Q35': 0, 'Q36': 0} | PASS |
| T58 | Excel – fórmulas | aba Qualidade: 36 testes, 36 PASS, 0 FAIL | 36/36/0 | 36/36/0 | PASS |
| T59 | Excel – fórmulas | aba Indicadores: conferências internas | 12 de 12 | 12 de 12 | PASS |
| T60 | Excel – resumos | resumo por mês = agrupamento independente dos CSVs | 0 | 0 | PASS |
| T61 | Excel – resumos | resumo por categoria = agrupamento independente | 0 | 0 | PASS |
| T62 | Excel – resumos | resumo por centro de trabalho = agrupamento independente | 0 | 0 | PASS |
| T63 | Excel – resumos | resumo por material = agrupamento independente (20 materiais) | 0 | 0 | PASS |
| T64 | Reconciliação de KPIs | KPI 'Unidades consumidas': Excel = cálculo independente dos CSVs | 1647 | 1647 | PASS |
| T65 | Reconciliação de KPIs | KPI 'Custo estimado consumido': Excel = cálculo independente dos CSVs | 24606.20 | 24606.2 | PASS |
| T66 | Reconciliação de KPIs | KPI 'Movimentações': Excel = cálculo independente dos CSVs | 180 | 180 | PASS |
| T67 | Reconciliação de KPIs | KPI 'Materiais distintos movimentados': Excel = cálculo independente dos CSVs | 20 | 20 | PASS |
| T68 | Reconciliação de KPIs | KPI 'Itens abaixo ou iguais ao mínimo': Excel = cálculo independente dos CSVs | 5 | 5 | PASS |
| T69 | Reconciliação de KPIs | KPI 'Itens críticos abaixo ou iguais ao mínimo': Excel = cálculo independente dos CSVs | 2 | 2 | PASS |
| T70 | Detecção de erros | dataset com erros injetados: Python detecta exatamente o esperado nos 36 testes | 36/36 | 36/36 | PASS |
| T71 | Detecção de erros | dataset com erros injetados: fórmulas do Excel detectam exatamente o esperado nos 36 testes | 36/36 | 36/36 | PASS |
| T72 | Detecção de erros | Excel marca FAIL exatamente nos testes com erro injetado | 33 | 33 | PASS |
| T73 | Detecção de erros | Status_Estoque/Critico_Em_Risco no Excel = Python, incluindo estoque negativo e limites (Atual = Mín, = Mín x 1,25, logo acima) | 22/22 | 22/22 | PASS |
| T74 | Detecção de erros | limites: MAT002 (30/30) = REPOR; MAT001 (50 = 40 x 1,25) = ATENÇÃO; MAT003 (13 > 12,5) = OK | REPOR/ATENÇÃO/OK | REPOR/ATENÇÃO/OK | PASS |
| T75 | Segurança | varredura de PII/credenciais em CSVs, XLSX, scripts e documentos (e-mail, telefone, CPF/CNPJ, senha/token/chave, nome de empresa real) | 0 | 0 | PASS |
| T76 | Segurança | sem .env, .pbix, chaves ou credenciais no pacote | 0 | 0 | PASS |
| T77 | Segurança | fornecedores fictícios ('Fornecedor …') (Q36) | 0 | 0 | PASS |
| T78 | Documentação | arquivos obrigatórios existem (o ZIP e o TEST_REPORT são conferidos à parte) | [] | [] | PASS |
| T79 | Documentação | estrutura de pastas do enunciado | [] | [] | PASS |
| T80 | Documentação | aviso 'Dados 100% sintéticos…' em README, DATA_DICTIONARY, LIMITATIONS, V0_1_NOTES | [] | [] | PASS |
| T81 | Documentação | CHANGELOG inicia com 'V0.1 — Foundation' | True | True | PASS |
| T82 | Documentação | DATA_DICTIONARY descreve todas as colunas dos 3 CSVs | [] | [] | PASS |
| T83 | Documentação | exemplos do DATA_DICTIONARY existem de fato nas colunas dos CSVs | [] | [] | PASS |
| T84 | Documentação | DATA_DICTIONARY inclui as 4 colunas calculadas do Excel | True | True | PASS |
| T85 | Documentação | README e V0_1_NOTES citam os 6 KPIs com os valores reais (tabela) | [] | [] | PASS |
| T86 | Documentação | LIMITATIONS cobre os 8 pontos exigidos | [] | [] | PASS |
| T87 | Documentação | V0_1_NOTES contém 'O que o autor precisa saber explicar' e os 4 blocos (O QUE FOI FEITO / POR QUE FOI FEITO / COMO VALIDAR / O QUE PRECISO SABER EXPLICAR) | True | True | PASS |
| T88 | Escopo | Projeto 1 sem forecasting: nenhuma biblioteca/modelo de previsão nos scripts e nenhuma função FORECAST/TREND/GROWTH nas fórmulas do Excel | [] | [] | PASS |
| T89 | Segurança | controle positivo: o detector de PII/segredos sinaliza 7 amostras plantadas (e-mail, telefone, CPF, CNPJ, senha, chave, empresa real) | [] | [] | PASS |
| T90 | Segurança | controle negativo: texto limpo do catálogo não gera alerta | 0 | 0 | PASS |
| T91 | Portabilidade | scripts sem caminhos absolutos do ambiente original e sem referência a script de recálculo externo ao projeto | [] | [] | PASS |
| T92 | Portabilidade | LOBO_SOFFICE inválido: localizar_soffice() = None e recalcular() levanta LibreOfficeIndisponivel com orientação (sem travar) | None / LibreOfficeIndisponivel | None / LibreOfficeIndisponivel | PASS |
| T93 | Portabilidade | LibreOffice localizado neste ambiente (variável LOBO_SOFFICE ou PATH) | encontrado | encontrado | PASS |
| T94 | Portabilidade | reconstruir o XLSX a partir dos CSVs (recálculo portátil) reproduz o XLSX entregue: fórmulas/constantes e valores calculados de todas as células | 0 / 0 diferenças | 0 / 0 diferenças em 3278 células (857 fórmulas, 0 erros) | PASS |
| T95 | Portabilidade | validar_dados.py executado a partir de outro diretório de trabalho: código de saída 0 e 36 PASS | 0 / 36 PASS / 0 FAIL | 0 / 36 PASS / 0 FAIL | PASS |
| T96 | Regressão vs V0.1 | arquivos que a V0.1.1 não pode alterar (CSVs, XLSX, regras em lobo_common/gerar_dados/validar_dados, DATA_DICTIONARY, V0_1_NOTES, relatório da V0.1) seguem byte a byte iguais à base congelada | [] | [] | PASS |
| T97 | Regressão vs V0.1 | a base congelada registra exatamente 10 arquivos protegidos | 10 | 10 | PASS |
| T98 | Documentação V0.1.1 | arquivos novos da V0.1.1 presentes | [] | [] | PASS |
| T99 | Documentação V0.1.1 | CHANGELOG registra 'V0.1.1 — Portable Baseline' acima da seção V0.1 (preservada); V0_1_1_NOTES traz o aviso de dados sintéticos | True | True | PASS |
| T100 | Portabilidade | requirements.txt declara pandas, openpyxl e lxml | [] | [] | PASS |
| T101 | Portabilidade | o gerador do relatório lê as notas manuais de dentro do pacote (não de pastas externas) | 2 | 2 | PASS |
| T102 | Qualidade do código | sem imports não usados nos scripts | [] | [] | PASS |
| T103 | Qualidade do código | sem f-strings que exijam Python >= 3.12 (barra invertida ou mesma aspa dentro do campo) | [] | [] | PASS |
| T104 | Regressão vs V0.1.1 | arquivos protegidos da V0.1.1 (CSVs, XLSX V0.1, scripts e regras, testes, DATA_DICTIONARY, notas e relatórios históricos) seguem byte a byte iguais | [] | [] | PASS |
| T105 | Regressão vs V0.1.1 | a base V0.1.1 registra 23 arquivos protegidos, incluindo os 3 CSVs, o XLSX histórico e a suíte executar_testes.py | 23 | 23 | PASS |
| T106 | Regressão vs V0.1.1 | os 3 CSVs oficiais têm os SHA-256 da baseline informados no enunciado da V0.2 (sem alteração de dados) | [] | [] | PASS |
| T107 | Regressão vs V0.1.1 | XLSX histórico V0.1.1 (lobo_insights_industrial_v0_1.xlsx) preservado e distinto do XLSX V0.2 | True | True | PASS |
| T108 | KPIs oficiais (regressão) | 'Unidades consumidas': referência 1647 = XLSX V0.1.1 = Indicadores V0.2 = cartão do DASHBOARD = CSV independente | 1647 | 1647 / 1647 / 1647 / 1647 | PASS |
| T109 | KPIs oficiais (regressão) | 'Custo estimado consumido': referência 24606.20 = XLSX V0.1.1 = Indicadores V0.2 = cartão do DASHBOARD = CSV independente | 24606.20 | 24606.2 / 24606.2 / 24606.2 / 24606.20 | PASS |
| T110 | KPIs oficiais (regressão) | 'Movimentações': referência 180 = XLSX V0.1.1 = Indicadores V0.2 = cartão do DASHBOARD = CSV independente | 180 | 180 / 180 / 180 / 180 | PASS |
| T111 | KPIs oficiais (regressão) | 'Materiais distintos movimentados': referência 20 = XLSX V0.1.1 = Indicadores V0.2 = cartão do DASHBOARD = CSV independente | 20 | 20 / 20 / 20 / 20 | PASS |
| T112 | KPIs oficiais (regressão) | 'Itens abaixo ou iguais ao mínimo': referência 5 = XLSX V0.1.1 = Indicadores V0.2 = cartão do DASHBOARD = CSV independente | 5 | 5 / 5 / 5 / 5 | PASS |
| T113 | KPIs oficiais (regressão) | 'Itens críticos abaixo ou iguais ao mínimo': referência 2 = XLSX V0.1.1 = Indicadores V0.2 = cartão do DASHBOARD = CSV independente | 2 | 2 / 2 / 2 / 2 | PASS |
| T114 | KPIs oficiais (regressão) | rótulos dos 6 cartões do DASHBOARD = nomes oficiais dos KPIs (Indicadores) | ['Unidades consumidas', 'Custo estimado consumido', 'Movimentações', 'Materiais distintos movimentados', 'Itens abaixo ou iguais ao mínimo', 'Itens críticos abaixo ou iguais ao mínimo'] | ['Unidades consumidas', 'Custo estimado consumido', 'Movimentações', 'Materiais distintos movimentados', 'Itens abaixo ou iguais ao mínimo', 'Itens críticos abaixo ou iguais ao mínimo'] | PASS |
| T115 | Excel V0.2 – estrutura | abas na ordem (DASHBOARD primeiro) | ['DASHBOARD', 'LEIA_ME', 'Indicadores', 'Analises', 'Materiais', 'Estoque', 'Consumo', 'Qualidade', 'Parametros'] | ['DASHBOARD', 'LEIA_ME', 'Indicadores', 'Analises', 'Materiais', 'Estoque', 'Consumo', 'Qualidade', 'Parametros'] | PASS |
| T116 | Excel V0.2 – estrutura | DASHBOARD é a aba ativa ao abrir (única selecionada) | DASHBOARD / [DASHBOARD] | DASHBOARD / ['DASHBOARD'] | PASS |
| T117 | Excel V0.2 – estrutura | tabelas estruturadas e intervalos idênticos aos da V0.1.1 | {'tbMateriais': 'A1:G21', 'tbEstoque': 'A1:L21', 'tbConsumo': 'A1:L181', 'tbQualidade': 'A9:E45'} | {'tbMateriais': 'A1:G21', 'tbEstoque': 'A1:L21', 'tbConsumo': 'A1:L181', 'tbQualidade': 'A9:E45'} | PASS |
| T118 | Excel V0.2 – estrutura | nomes definidos = 9 da V0.1.1 + pLimitePareto, pTopA, pTopB | ['lstCategorias', 'lstCentros', 'lstCriticidade', 'lstMeses', 'lstTipoMov', 'lstUnidades', 'pFimPeriodo', 'pInicioPeriodo', 'pLimitePareto', 'pMargemAtencao', 'pTopA', 'pTopB'] | ['lstCategorias', 'lstCentros', 'lstCriticidade', 'lstMeses', 'lstTipoMov', 'lstUnidades', 'pFimPeriodo', 'pInicioPeriodo', 'pLimitePareto', 'pMargemAtencao', 'pTopA', 'pTopB'] | PASS |
| T119 | Excel V0.2 – estrutura | sem vínculos externos (externalLinks) e sem relações externas (TargetMode=External) | 0 / 0 | 0 / 0 | PASS |
| T120 | Excel V0.2 – estrutura | recálculo total ao abrir (fullCalcOnLoad) | True | True | PASS |
| T121 | Excel V0.2 – estrutura | nenhum caminho absoluto/privado dentro do XLSX (fórmulas, gráficos, hiperlinks, propriedades) | [] | [] | PASS |
| T122 | Excel V0.2 – estrutura | 6 hiperlinks do DASHBOARD são internos (location) e nenhum aponta para fora do arquivo | 6 / 0 externos | 6 / 0 externos | PASS |
| T123 | Excel V0.2 – estrutura | destinos dos hiperlinks existem (abas do próprio arquivo) | [] | [] | PASS |
| T124 | Preservação da base | aba Materiais: todas as células (constantes e fórmulas) idênticas à V0.1.1; validações de dados e formatação condicional iguais | 0 diferenças | 0 diferenças; DV 4/4; CF 1/1 | PASS |
| T125 | Preservação da base | aba Estoque: todas as células (constantes e fórmulas) idênticas à V0.1.1; validações de dados e formatação condicional iguais | 0 diferenças | 0 diferenças; DV 6/6; CF 3/3 | PASS |
| T126 | Preservação da base | aba Consumo: todas as células (constantes e fórmulas) idênticas à V0.1.1; validações de dados e formatação condicional iguais | 0 diferenças | 0 diferenças; DV 8/8; CF 0/0 | PASS |
| T127 | Preservação da base | aba Qualidade: todas as células (constantes e fórmulas) idênticas à V0.1.1; validações de dados e formatação condicional iguais | 0 diferenças | 0 diferenças; DV 0/0; CF 1/1 | PASS |
| T128 | Preservação da base | aba Indicadores: fórmulas idênticas à V0.1.1; únicas diferenças = título A1 e nota A3 (texto) | ['A1', 'A3'] | ['A1', 'A3'] | PASS |
| T129 | Preservação da base | aba Parametros: A1:F18 idênticas à V0.1.1; todas as diferenças estão em A20:C24 (Limite_Pareto=0,8; Top_N_A=3; Top_N_B=5) | 0 diferenças fora de A20:C24 / [0.8, 3, 5] | 0 diferenças / [0.8, 3, 5] | PASS |
| T130 | Preservação da base | valores calculados de Materiais/Estoque/Consumo/Qualidade/Indicadores = V0.1.1 (todas as células) | 0 | 0 | PASS |
| T131 | Excel V0.2 – fórmulas | fórmulas reais no arquivo (> 2000; V0.1.1 tinha 857) | > 2000 | 2417 | PASS |
| T132 | Excel V0.2 – fórmulas | células com erro (#REF!, #DIV/0!, #VALUE!, #NAME?, #N/A…) nos valores calculados: todas as abas | 0 | 0 | PASS |
| T133 | Excel V0.2 – fórmulas | nenhum '#REF!' ou intervalo quebrado dentro do XML (fórmulas, nomes definidos, gráficos) | 0 | 0 | PASS |
| T134 | Excel V0.2 – fórmulas | toda célula de fórmula tem valor em cache gravado (<v>) — visualizadores e celular mostram os números | 0 | 0 | PASS |
| T135 | Excel V0.2 – fórmulas | nenhum número digitado (constante numérica) nas abas DASHBOARD, Analises e Indicadores — tudo é fórmula | 0 / 0 / 0 | 0 / 0 / 0 | PASS |
| T136 | Escopo | nenhuma função de previsão/tendência/regressão (FORECAST, TREND, GROWTH, LINEST, LOGEST, SLOPE, INTERCEPT) nas fórmulas do XLSX V0.2 | [] | [] | PASS |
| T137 | Compatibilidade | nenhuma função de matriz dinâmica/pós-2019 (XLOOKUP, SORT, FILTER, UNIQUE, SEQUENCE, LET, LAMBDA) — ordenações por RANK+COUNTIF+INDEX/MATCH | [] | [] | PASS |
| T138 | Auditabilidade | fórmulas do DASHBOARD só leem as abas Analises, Indicadores e Qualidade (cadeia dados → Indicadores → Analises → DASHBOARD) | ['Analises', 'Indicadores', 'Qualidade'] | ['Analises', 'Indicadores', 'Qualidade'] | PASS |
| T139 | Auditabilidade | fórmulas da aba Analises só leem Indicadores, Parametros, tabelas estruturadas e o DASHBOARD (apenas na conferência dos cartões) | ['DASHBOARD', 'Indicadores', 'Parametros'] | ['DASHBOARD', 'Indicadores', 'Parametros'] | PASS |
| T140 | Análises – valores | TODAS as tabelas da aba Analises (mensal, categoria, centro, Pareto de custo e de consumo, concentração, fila de estoque, painel, matriz, conferências) = cálculo independente dos CSVs (Decimal) | 0 divergências | 0 divergências | PASS |
| T141 | Análises – valores | resultado esperado dos dados oficiais: 9 materiais somam 80% do custo; 4 materiais somam 80% das unidades | 9 / 4 | 9 / 4 | PASS |
| T142 | Análises – valores | há empates reais nas unidades por material (MAT011/MAT012/MAT020 = 11; MAT006/MAT018 = 5) e o desempate é a ordem do catálogo | MAT011,MAT012,MAT020 / MAT006,MAT018 | MAT011,MAT012,MAT020 / MAT006,MAT018 | PASS |
| T143 | Análises – valores | Pareto: % acumulado é não decrescente e termina em 100%; soma dos valores = KPI (custo e unidades) | True | True | PASS |
| T144 | Análises – valores | conferências internas da aba Analises (seção 8) e de Indicadores (seção 6) | 18 de 18 / 12 de 12 | 18 de 18 / 12 de 12 | PASS |
| T145 | DASHBOARD – valores | cartões de KPI e fila de atenção do DASHBOARD = cálculo independente dos CSVs | 0 divergências | 0 divergências | PASS |
| T146 | DASHBOARD – valores | textos de leitura gerados por fórmula trazem os fatos esperados dos dados (mês de maior consumo, maiores categorias/centros, concentração, contagens de estoque, conferências) | [] | [] | PASS |
| T147 | DASHBOARD – valores | variações mensais nas leituras: +13,7% (jul/jun) e -49,6% (ago/jul) para unidades; +29,3% e -52,3% para custo (separador decimal do LibreOffice ou do Excel) | [] | [] | PASS |
| T148 | Escopo | DASHBOARD declara: dados sintéticos, sem previsão e sem recomendação operacional | [] | [] | PASS |
| T149 | Escopo | DASHBOARD não contém linguagem de previsão/projeção (fora das negações explícitas) | [] | [] | PASS |
| T150 | Gráficos | 9 gráficos no DASHBOARD (6 de barras simples, 2 Pareto combinados, 1 comparativo de estoque) e nenhum em outra aba | 9 / 0 | 9 / 0 | PASS |
| T151 | Gráficos | cada gráfico: tipo, orientação, nº de séries, nº de pontos e intervalos (valores e categorias) = o esperado da camada Analises | [] | [] | PASS |
| T152 | Gráficos | todas as referências apontam só para a aba Analises, em células preenchidas e numéricas (nenhum #REF!, nenhuma célula vazia) | [] | [] | PASS |
| T153 | Gráficos | todos os gráficos declaram 'sem título automático' e têm rótulos de dados; Paretos têm eixo secundário (2 eixos de valor) | True | True | PASS |
| T154 | Gráficos | os 9 gráficos estão ancorados por células (twoCellAnchor) dentro da área impressa do DASHBOARD | 9 | 9 | PASS |
| T155 | Layout / impressão | DASHBOARD: paisagem, A4, largura ajustada a 1 página, área de impressão definida, 2 quebras manuais de página, sem linhas de grade | landscape/9/1/2/sem grade | landscape/9/1/2/sem grade | PASS |
| T156 | Layout / impressão | DASHBOARD: 12 colunas de mesma largura (grade uniforme) e sem congelamento de painéis (decisão: rolagem livre em notebook) | 12 x 13 / sem freeze | {13.0} / None | PASS |
| T157 | Layout / impressão | abas de dados congelam o cabeçalho (Materiais/Estoque/Consumo A2; Qualidade A10; Indicadores A4; Analises A6) | A2/A2/A2/A10/A4/A6 | A2/A2/A2/A10/A4/A6 | PASS |
| T158 | Layout / impressão | Analises: paisagem e largura ajustada; Analises sem linhas de grade | True | True | PASS |
| T159 | Layout / impressão | DASHBOARD: formatação condicional nos status de estoque, criticidade, matriz e cartões de alerta (>= 6 intervalos) | >= 6 | 8 | PASS |
| T160 | Layout / impressão | DASHBOARD usa uma única família tipográfica (Arial) | {'Arial'} | {'Arial'} | PASS |
| T161 | Layout / impressão | cartões de KPI: linhas de rótulo, valor e leitura com altura definida (27/38/28 pt) para o texto não ser cortado | [27, 38, 28] | [27.0, 38.0, 28.0] | PASS |
| T162 | Layout / impressão | cartões de KPI: formatos numéricos (unidades #,##0; custo #,##0.00; contagens #,##0) | ['#,##0', '#,##0.00', '#,##0', '#,##0', '#,##0', '#,##0'] | ['#,##0', '#,##0.00', '#,##0', '#,##0', '#,##0', '#,##0'] | PASS |
| T163 | Reconstrução / mutação | reconstruir o XLSX V0.2 a partir dos CSVs (construir_excel_v0_2.py + recálculo) reproduz o XLSX entregue: fórmulas/constantes e valores de todas as células | 0 / 0 diferenças | 0 / 0 diferenças em 5074 células (2417 fórmulas, 0 erros) | PASS |
| T164 | Reconstrução / mutação | teste de mutação: dados alterados (consumo de MAT010 vira 'Ajuste', MAT019 = 300 un., MAT004 muda de mês, estoques de MAT001/005/008/016 alterados) => TODAS as tabelas da Analises e os cartões/fila do DASHBOARD recalculam e batem com o cálculo independente dos dados alterados | 0 / 0 divergências (0 erros) | 0 / 0 divergências (0 erros) | PASS |
| T165 | Reconstrução / mutação | a mutação é efetiva: os resultados esperados MUDAM em relação aos dados oficiais (KPI 4 = 19, ranking, REPOR, overflow do painel) | >= 6 indicadores diferentes | 8 de 8 diferentes | PASS |
| T166 | Inspeção visual (automatizada) | PDF do workbook gerado; as 3 primeiras páginas são o DASHBOARD (cabeçalho e KPIs; Pareto; estoque) | LOBO INSIGHTS / Pareto de custo / Fila de atenção | LOBO INSIGHTS / Pareto de custo / Fila de atenção | PASS |
| T167 | Inspeção visual (automatizada) | gráficos realmente desenhados: pixels das cores de série (azul de unidades e cobre de custo) presentes nas páginas 1 e 2 e barras de estoque na 3 | azul>2000 e cobre>2000 (p1 e p2); azul>500 (p3) | p1 (10651, 14522); p2 (8355, 10034); p3 (2682, 2) | PASS |
| T168 | Documentação V0.2 | arquivos novos da V0.2 presentes | [] | [] | PASS |
| T169 | Documentação V0.2 | V0_2_NOTES.md contém aviso de dados sintéticos e as 7 seções exigidas | [] | [] | PASS |
| T170 | Documentação V0.2 | CHANGELOG registra 'V0.2 — Analytics & Dashboard' acima de V0.1.1 e V0.1 (históricos preservados) | True | True | PASS |
| T171 | Documentação V0.2 | README cita 'V0.2', o XLSX V0.2, o DASHBOARD e os 6 KPIs com os valores reais | [] | [] | PASS |
| T172 | Documentação V0.2 | LIMITATIONS declara: sintético, sem previsão, sem recomendação operacional e as limitações novas da V0.2 (Excel real, filtros, cache/pré-visualização) | [] | [] | PASS |
| T173 | Documentação V0.2 | docs/TEST_REPORT_V0_1_1.md é o relatório histórico da V0.1.1 (título e 103 testes) | True | True | PASS |
| T174 | Documentação V0.2 | a suíte legada foi executada sem alteração: executar_testes.py idêntico ao da V0.1.1 e 103 testes legados registrados antes dos novos | idêntico / 103 | idêntico / 103 | PASS |
| T175 | Nomenclatura (V0.2.1) | o XLSX atual se chama EXATAMENTE Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx e existe em excel/ | Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx | Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx | PASS |
| T176 | Nomenclatura (V0.2.1) | o pacote da release se chama EXATAMENTE Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip (nome esperado, constante do módulo de release) | Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip | Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip | PASS |
| T177 | Nomenclatura (V0.2.1) | XLSX e ZIP começam com 'Projeto_Lobo_1_' e têm o mesmo radical (só muda a extensão) | True | True | PASS |
| T178 | Nomenclatura (V0.2.1) | a regra de validação aceita os 2 nomes obrigatórios e REJEITA nomes ambíguos (Lobo_…, lobo_…, projeto/produto 2, extensão errada) | 2 aceitos / 0 dos 7 ambíguos aceitos | 2 aceitos / 0 dos 7 ambíguos aceitos | PASS |
| T179 | Nomenclatura (V0.2.1) | os entregáveis principais desta versão (XLSX atual e ZIP) não casam com nenhum padrão ambíguo | [] | [] | PASS |
| T180 | Nomenclatura (V0.2.1) | excel/ contém só o XLSX atual e os 2 nomes históricos (V0.1.1 e V0.2, não renomeados); nenhum outro .xlsx no projeto | ['excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx', 'excel/lobo_insights_industrial_v0_1.xlsx', 'excel/lobo_insights_industrial_v0_2.xlsx'] | ['excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx', 'excel/lobo_insights_industrial_v0_1.xlsx', 'excel/lobo_insights_industrial_v0_2.xlsx'] | PASS |
| T181 | Nomenclatura (V0.2.1) | os nomes históricos da V0.2 e anteriores foram preservados (XLSX V0.1.1, XLSX V0.2, scripts e documentos das versões anteriores) | [] | [] | PASS |
| T182 | Nomenclatura (V0.2.1) | a constante de geração do XLSX (construir_excel_v0_2_1.py) aponta para o nome obrigatório e main() grava nele | constante correta / main usa NOME_XLSX | Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx / usa | PASS |
| T183 | Nomenclatura (V0.2.1) | README, CHANGELOG e V0_2_1_NOTES citam os dois nomes obrigatórios exatos | [] | [] | PASS |
| T184 | Regressão vs V0.2 auditada | o XLSX atual é byte a byte igual ao XLSX auditado da V0.2 (mesmo SHA-256) ⇒ nenhuma alteração de dados, fórmulas, KPIs, gráficos, layout ou resultados | 81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846 | 81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846 / 81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846 | PASS |
| T185 | Regressão vs V0.2 auditada | 30 arquivos da V0.2 auditada (CSVs, XLSX históricos, todos os scripts e testes da V0.1.1/V0.2, documentos históricos, notas da V0.2) seguem byte a byte iguais | 30 / 0 diferenças | 30 / 0 diferenças | PASS |
| T186 | Regressão vs V0.2 auditada | a linha de base da V0.2 inclui dados, regras, XLSX históricos e todo o código analítico e de teste (nada foi excluído da proteção) | [] | [] | PASS |
| T187 | Regressão vs V0.2 auditada | o gerador da V0.2.1 (construir_excel_v0_2_1.gerar) reproduz o XLSX atual: mesmas fórmulas/constantes, mesmos valores e mesmos 9 gráficos em todas as células (bytes diferem, conteúdo não) | 0 / 0 diferenças / 9 e 9 gráficos / 0 erros | 0 / 0 diferenças em 5074 células / 9 e 9 gráficos / 0 erros | PASS |
| T188 | Manifesto SHA-256 (V0.2.1) | SHA256SUMS.txt segue o formato do 'sha256sum -c' (sem comentários), está em ordem, sem duplicatas e sem listar a si mesmo, o snapshot nem os regeneráveis | True | True | PASS |
| T189 | Manifesto SHA-256 (V0.2.1) | SHA256SUMS.txt confere integralmente (todos os arquivos estáveis existem e têm o hash registrado) | 0 ausentes / 0 divergentes | 0 ausentes / 0 divergentes de 40 | PASS |
| T190 | Manifesto SHA-256 (V0.2.1) | a ferramenta do sistema 'sha256sum -c SHA256SUMS.txt' termina com sucesso | código 0 | código 0; 40 OK;  | PASS |
| T191 | Manifesto SHA-256 (V0.2.1) | os regeneráveis são exatamente 4 (relatório + 3 prévias), estão fora do manifesto operacional e seguem declarados no módulo de release | ['docs/TEST_REPORT.md', 'images/dashboard_v0_2_pagina_1.png', 'images/dashboard_v0_2_pagina_2.png', 'images/dashboard_v0_2_pagina_3.png'] | ['docs/TEST_REPORT.md', 'images/dashboard_v0_2_pagina_1.png', 'images/dashboard_v0_2_pagina_2.png', 'images/dashboard_v0_2_pagina_3.png'] | PASS |
| T192 | Manifesto SHA-256 (V0.2.1) | nenhum arquivo do projeto ficou fora do controle: tudo o que não é regenerável nem manifesto está em SHA256SUMS.txt (sem exceções silenciosas) | [] | [] | PASS |
| T193 | Manifesto SHA-256 (V0.2.1) | SHA256SUMS.txt é reproduzível: recalculado a partir da árvore atual (gerar_manifesto_estavel) é idêntico ao arquivo gravado | True | True | PASS |
| T194 | Manifesto SHA-256 (V0.2.1) | RELEASE_SHA256SUMS.txt (snapshot) lista TODOS os arquivos exceto ele mesmo, inclusive SHA256SUMS.txt e os 4 regeneráveis, no formato sha256sum | True | True | PASS |
| T195 | Manifesto SHA-256 (V0.2.1) | snapshot: nenhum arquivo ausente e nenhuma divergência FORA dos regeneráveis (só os 4 regeneráveis podem diferir do instante do empacotamento; no pacote recém-extraído: 0) | 0 ausentes / divergências ⊆ regeneráveis | 0 ausentes / divergentes: [] | PASS |
| T196 | Manifesto SHA-256 (V0.2.1) | controle (cópia descartável): execução legítima que muda os 4 regeneráveis — ou apaga um deles — NÃO invalida SHA256SUMS.txt, mas o snapshot registra a mudança | 3 / 3 verificações | 3 / 3 verificações | PASS |
| T197 | Manifesto SHA-256 (V0.2.1) | controle (cópia descartável): alteração REAL num CSV, num script, no XLSX ou numa nota, arquivo estável apagado e arquivo novo não registrado são detectados (nada é mascarado) | 6 / 6 detecções | 6 / 6 detecções | PASS |
| T198 | Documentação V0.2.1 | CHANGELOG registra 'V0.2.1 — Packaging & Reproducibility Fix' acima de V0.2, V0.1.1 e V0.1 (históricos preservados) | True | True | PASS |
| T199 | Documentação V0.2.1 | V0_2_1_NOTES.md contém o aviso de dados sintéticos, as 8 seções exigidas, os dois manifestos e os 4 arquivos regeneráveis | [] | [] | PASS |
| T200 | Documentação V0.2.1 | README cita V0.2.1, os dois manifestos, o executor da V0.2.1 e continua com a tabela dos 6 KPIs e o histórico (XLSX V0.1/V0.2) | [] | [] | PASS |
| T201 | Documentação V0.2.1 | docs/manual/observacoes_execucao_v0_2_1.md existe e registra o que foi (e não foi) reproduzido do problema dos hashes | True | True | PASS |

## 6. Regressão dos 6 KPIs oficiais
| KPI | Referência | XLSX V0.1.1 | Indicadores (XLSX atual) | Cartão do DASHBOARD | CSV independente | Status |
|---|---|---|---|---|---|---|
| Unidades consumidas | 1647 | 1647 | 1647 | 1647 | 1647 | PASS |
| Custo estimado consumido | 24606.20 | 24606.2 | 24606.2 | 24606.2 | 24606.20 | PASS |
| Movimentações | 180 | 180 | 180 | 180 | 180 | PASS |
| Materiais distintos movimentados | 20 | 20 | 20 | 20 | 20 | PASS |
| Itens abaixo ou iguais ao mínimo | 5 | 5 | 5 | 5 | 5 | PASS |
| Itens críticos abaixo ou iguais ao mínimo | 2 | 2 | 2 | 2 | 2 | PASS |

## 7. Inventário dos gráficos e teste de mutação (executados sobre o XLSX atual)
| Arquivo | Tipo | Direção | Séries | Pontos | Rótulos | Referências |
|---|---|---|---|---|---|---|
| chart1.xml | barChart | col | 1 | 3 | 1 | 'Analises'!$B$8:$B$10; 'Analises'!$D$8:$D$10 |
| chart2.xml | barChart | col | 1 | 3 | 1 | 'Analises'!$B$8:$B$10; 'Analises'!$E$8:$E$10 |
| chart3.xml | barChart | bar | 1 | 6 | 1 | 'Analises'!$B$26:$B$31; 'Analises'!$C$26:$C$31 |
| chart4.xml | barChart | bar | 1 | 6 | 1 | 'Analises'!$G$26:$G$31; 'Analises'!$H$26:$H$31 |
| chart5.xml | barChart | bar | 1 | 5 | 1 | 'Analises'!$B$45:$B$49; 'Analises'!$C$45:$C$49 |
| chart6.xml | barChart | bar | 1 | 5 | 1 | 'Analises'!$G$45:$G$49; 'Analises'!$H$45:$H$49 |
| chart7.xml | barChart+lineChart | col | 4 | 20 | 0 | 'Analises'!I78; 'Analises'!$C$79:$C$98; 'Analises'!$I$79:$I$98; 'Analises'!J78; 'Analises'!$C$79:$C$98; 'Analises'!$J$79:$J$98; 'Analises'!F78; 'Analises'!$C$79:$C$98; 'Analises'!$F$79:$F$98; 'Analises'!G78; 'Analises'!$C$79:$C$98; 'Analises'!$G$79:$G$98 |
| chart8.xml | barChart+lineChart | col | 4 | 20 | 0 | 'Analises'!I103; 'Analises'!$C$104:$C$123; 'Analises'!$I$104:$I$123; 'Analises'!J103; 'Analises'!$C$104:$C$123; 'Analises'!$J$104:$J$123; 'Analises'!F103; 'Analises'!$C$104:$C$123; 'Analises'!$F$104:$F$123; 'Analises'!G103; 'Analises'!$C$104:$C$123; 'Analises'!$G$104:$G$123 |
| chart9.xml | barChart | bar | 2 | 10 | 2 | 'Analises'!D200; 'Analises'!$B$201:$B$210; 'Analises'!$D$201:$D$210; 'Analises'!E200; 'Analises'!$B$201:$B$210; 'Analises'!$E$201:$E$210 |

| Indicador (mutação) | Dados oficiais | Esperado com dados alterados | Obtido no Excel recalculado |
|---|---|---|---|
| KPI Materiais distintos movimentados | 20 | 19 | 19 |
| KPI Itens abaixo ou iguais ao mínimo (REPOR) | 5 | 6 | 6 |
| KPI Itens críticos (Alta em REPOR) | 2 | 1 | 1 |
| KPI Unidades consumidas | 1647 | 1943 | 1943 |
| Materiais que somam 80% das unidades | 4 | 5 | 5 |
| 2º do ranking de unidades (material) | MAT007 | MAT019 | MAT019 |
| Itens fora do OK (estoque) | 9 | 11 | 11 |
| Itens não exibidos no painel | 0 | 1 | 1 |

## 8. Inspeção visual (da V0.2; sem alteração visual na V0.2.1)
**Método.** O XLSX V0.2 foi convertido em PDF pelo LibreOffice (modo headless), as páginas foram rasterizadas com `pdftoppm` e **as imagens foram abertas e examinadas visualmente pelo executor (Claude)**. Isso mostra como o LibreOffice desenha o arquivo; **não** é uma inspeção no Microsoft Excel.

**Páginas inspecionadas (com resultado):**

| Aba / página | Resultado da inspeção |
|---|---|
| DASHBOARD, página 1 (cabeçalho, KPIs, evolução mensal, categorias) | Faixa superior em azul-marinho com filete cobre; 6 cartões legíveis, números em destaque, leitura rápida sob cada valor; gráficos de colunas (mensal) e barras horizontais (categorias) com rótulos de dados; sem sobreposição de texto. Ajustado durante a inspeção: altura do rótulo dos cartões (títulos de 2 linhas estavam cortados) e margem à direita das barras (o rótulo do maior valor encostava na borda). |
| DASHBOARD, página 2 (centros, Pareto de custo e de consumo, cartões de concentração) | Barras dos centros legíveis; os dois Paretos mostram colunas do "núcleo" em cor forte, "demais" em tom claro, linha de % acumulado com marcadores e linha tracejada do limite; eixo secundário 0–100%. Nomes dos 20 materiais rotacionados em 90°, legíveis. Ajustado: altura do gráfico (a área de plotagem estava baixa) e quebra de página (a página 2 ficou só com "Centros" na primeira tentativa). |
| DASHBOARD, página 3 (estoque e criticidade) | Fila de atenção com REPOR em vermelho-claro e ATENÇÃO em âmbar; itens críticos em risco (Alta em REPOR) com texto vermelho; matriz criticidade × status com cores; gráfico Atual × Mínimo legível; rodapé com rastreabilidade, conferências e aviso de dados sintéticos. Ajustado: título da matriz estava truncado. |
| LEIA_ME | Texto legível; camadas, decisões (sem filtros) e SHA-256 dos CSVs presentes. |
| Analises, página 1 | Tabelas com cabeçalho azul-marinho, legíveis. Na impressão a página fica pequena (a aba é de cálculo e foi pensada para a tela). |
| Indicadores, página 1 | Idêntica à V0.1.1, exceto o título e a nota das linhas 1 e 3. |
| Parametros | Bloco novo (A20:C24) no mesmo padrão visual da V0.1; a primeira versão (colunas H:J) colidia com o texto longo da coluna C e foi movida. |

**NÃO inspecionado:** abas Consumo (páginas 13 a 16 do PDF), Qualidade, Materiais e Estoque (dados/tabelas idênticos aos da V0.1.1, comparados célula a célula por teste automatizado, mas não revistos visualmente nesta versão); páginas 8 em diante da aba Analises; **qualquer visualização no Microsoft Excel real**.

**Limites do método.** O LibreOffice difere do Excel em fontes, espaçamento interno dos gráficos, posição de rótulos e tratamento do eixo secundário. As proporções dos gráficos e a quebra de linha dos textos de leitura podem parecer diferentes no Excel. Além da inspeção visual, há verificações automatizadas de renderização (testes de "Inspeção visual (automatizada)" no relatório: texto de cada página do PDF e contagem de pixels das cores das séries).

**Prévias salvas:** `images/dashboard_v0_2_pagina_1.png`, `_2.png` e `_3.png` (geradas pela suíte de testes a partir do XLSX entregue).


## 9. Observações de execução
### V0.2
- **Base utilizada:** `Lobo_Insights_Industrial_V0_1_1.zip`, SHA-256 `97f4dc11fb16cbc6c18259af8556604c7d31b296c0c34e771dfb780850c6fcc4` (igual ao informado no enunciado). Os 27 arquivos conferiram com o `SHA256SUMS.txt` da V0.1.1.
- **Estado antes de qualquer alteração:** a suíte da V0.1.1, executada neste ambiente, deu 103 PASS / 0 FAIL / 0 SKIP (36/36 verificações de qualidade; 36/36 testes de detecção). Os 6 KPIs reproduziram 1.647 / 24.606,20 / 180 / 20 / 5 / 2. Nenhuma divergência a investigar.
- **Incidente de ambiente:** durante a construção, o diretório de trabalho do ambiente de execução foi reiniciado e a cópia de trabalho foi perdida. A baseline foi extraída de novo do ZIP original, o hash foi conferido outra vez e o trabalho recomeçou a partir dela. Nenhum arquivo da baseline foi editado nesse processo.
- **Defeito encontrado e corrigido (1) — gráficos vazios:** a primeira versão criava uma subclasse de `BarChart` para declarar `autoTitleDeleted`. Subclasses de classes `Serialisable` do openpyxl não herdam `__elements__`, e todos os gráficos saíam com `<barChart/>` vazio (sem séries). Foi percebido na inspeção visual (gráficos em branco). Correção: o ajuste passou a ser feito na instância (interceptando `_write()`), sem subclasse.
- **Defeito encontrado e corrigido (2) — título automático no Excel:** o primeiro ajuste procurava a tag com namespace, mas o openpyxl serializa sem prefixo; portanto `autoTitleDeleted` não era gravado. O LibreOffice não mostra título automático, então a inspeção visual **não** revelou o problema; ele apareceu ao inspecionar o XML e ao conferir os testes. No Excel, gráficos de uma série sem esse elemento exibem o nome da série como título. Correção aplicada e teste específico adicionado (todos os 9 gráficos declaram `autoTitleDeleted`).
- **Defeito encontrado e corrigido (3) — estilo perdido em células mescladas:** o openpyxl recria as células não-âncora de uma mesclagem sem estilo; bordas dos cartões precisam ser aplicadas depois de mesclar. Código reorganizado (`cartao()` e `bordas()`).
- **Ajustes visuais feitos após inspeção:** ver `inspecao_visual_v0_2.md`.
- **Controles positivos dos testes (verificação manual, cópia descartada):** defeitos injetados um por vez numa cópia do pacote e a suíte reprovou o teste correspondente em todos os casos — (A) valor em cache adulterado na aba Analises; (B) cartão do DASHBOARD trocado por constante; (C) um caractere alterado em um CSV; (D) gráfico sem `autoTitleDeleted`; (E) referência de gráfico quebrada (`#REF!`). No caso E a primeira versão do teste lançou exceção em vez de registrar FAIL; o teste foi tornado robusto e reexecutado.
- **Tentativas descartadas:** primeira posição dos parâmetros novos (colunas H:J da aba Parametros) — movidos para A20:C24; quebra de página após "Centros" — trocada por quebra antes.
- **Nenhum dado, KPI, regra, fórmula da V0.1.1 ou documento histórico foi alterado.** A aba Indicadores mudou apenas o texto das células A1 e A3 (título e nota) no XLSX novo; o XLSX histórico V0.1.1 permanece intocado e com o mesmo hash.


### V0.2.1
- **Pacote recebido:** `Lobo_Insights_Industrial_V0_2.zip` com SHA-256 `6c2007cb5d3658bb7d95ac335a1d91e8ba1b13e120a421db70f67faea5fadea4` (o mesmo da V0.2 entregue), 38 arquivos, `SHA256SUMS.txt` conferindo integralmente em cópia limpa. A suíte da V0.2 não faz referência ao manifesto (nenhum teste legado usa `SHA256SUMS.txt`).
- **Tentativa de reprodução do problema dos hashes — NÃO reproduzido.** Numa cópia do pacote recebido, `python scripts/executar_testes_v0_2.py` (modo normal) terminou com 174 PASS e `sha256sum -c SHA256SUMS.txt` **continuou passando**: neste ambiente (Python 3.12.3, LibreOffice 24.2.7.2, Linux x86_64) o relatório e as três imagens são regenerados com os mesmos bytes. Comparando a árvore antes e depois, a execução normal reescreve **exatamente 4 arquivos** (`docs/TEST_REPORT.md` e as 3 imagens) e nenhum outro; não cria nem remove arquivos. A causa relatada pela auditoria é **plausível e consistente** com o que se observa — o relatório grava a linha de ambiente (versões de Python, pandas, openpyxl, LibreOffice e plataforma) e as imagens são desenhadas pelo LibreOffice —, mas a variação de bytes em outro ambiente **não foi observada aqui** (é inferência, não medição).
- **Decisões de desenho:**
  - **XLSX = cópia byte a byte do auditado** (mesmo SHA-256), sob o nome novo, em vez de reconstruir: é a prova mais forte de "nenhuma alteração analítica ou visual" (reconstruir mudaria os bytes, pois os metadados do formato variam). Efeito colateral aceito e documentado: o texto interno do workbook continua dizendo "V0.2".
  - **XLSX histórico da V0.2 mantido** com o nome original (regra de não renomear nomes históricos; `docs/V0_2_NOTES.md` e a bateria da V0.2 continuam verdadeiros). O custo é um arquivo duplicado (~85 KB), garantido idêntico por teste.
  - **Nenhum arquivo da V0.2 foi editado** (`construir_excel_v0_2.py`, `executar_testes_v0_2.py` etc.); a geração no nome novo é um invólucro (`construir_excel_v0_2_1.py`).
  - **Pasta-raiz do ZIP** mantida como `lobo-insights-industrial/` (identidade do projeto entre versões).
  - **Regeneráveis = lista fechada de 4 caminhos**, com teste que exige exatamente essa lista e teste de cobertura para todos os demais arquivos.
- **Falhas encontradas nos meus próprios testes durante a construção (todas corrigidas):**
  1. O README ainda não citava o nome exato do ZIP: o teste de nomenclatura reprovou corretamente (T183) e o README foi corrigido.
  2. A varredura de escopo da suíte legada proíbe, em qualquer script, a palavra de previsão do Projeto 2; o módulo de release citava o prefixo ambíguo literalmente e passou a montá-lo por concatenação (verificado antes de rodar a suíte).
  3. Ao simular a remoção do XLSX atual, o executor **lançou exceção** em vez de registrar FAIL (`FileNotFoundError`, em 3 pontos). Tornado robusto e reexecutado: agora reprova com FAIL.
- **Detecção de manifesto velho (comportamento correto, não defeito):** durante o desenvolvimento editei README, CHANGELOG, notas e o executor depois de gerar os manifestos, e a suíte reprovou 6 testes de manifesto apontando exatamente esses 4 arquivos. Regerar os manifestos (`lobo_release.py manifestos`) resolveu. Ordem correta de release: editar → `manifestos` → executar testes → `manifestos` de novo (o snapshot passa a incluir os regeneráveis recém-gerados).
- **Controles positivos dos testes novos** (cópia descartável, um defeito por vez): (A) XLSX renomeado para nome ambíguo; (B) manifesto operacional incluindo um regenerável; (C) um CSV alterado; (D) `construir_excel_v0_2.py` alterado; (E) XLSX atual diferente do auditado; (F) arquivo novo não registrado; (G) nome-base do ZIP errado no módulo de release. **Todos reprovaram os testes correspondentes.**
- **Procedimento de release executado** (em ensaio com o mesmo conteúdo, repetido sobre o ZIP final):
  - Pacote recém-extraído: `sha256sum -c SHA256SUMS.txt` e `sha256sum -c RELEASE_SHA256SUMS.txt` — 0 divergências.
  - Suíte V0.2.1 em modo normal dentro do pacote extraído (regenera os 4 arquivos), depois `sha256sum -c SHA256SUMS.txt` — **código 0**. Nesse momento o snapshot acusou apenas os arquivos regeneráveis, como projetado (no ensaio, `docs/TEST_REPORT.md`, porque o relatório empacotado era de uma rodada anterior).
  - Suíte a partir de outro diretório (`cd /` e caminho absoluto): 201 testes PASS.
  - **Comando exato do auditor** (`python scripts/executar_testes_v0_2.py`, modo normal) num pacote extraído: 174 PASS e `sha256sum -c SHA256SUMS.txt` com código 0.
  - **Variação de ambiente simulada** (4 regeneráveis alterados de propósito): `SHA256SUMS.txt` continua passando (código 0) e o snapshot acusa exatamente os 4; **1 espaço acrescentado a um CSV**: `SHA256SUMS.txt` reprova (código 1) apontando só esse CSV.
- **O que este registro não garante:** o SHA-256 do ZIP final não está aqui (um arquivo não contém a prova do próprio ZIP); comportamento de bytes de PNG/relatório em ambientes diferentes do desta execução; Microsoft Excel real (mesmas limitações da V0.2).

