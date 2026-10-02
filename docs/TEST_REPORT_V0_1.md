# TEST_REPORT — Lobo Insights Industrial V0.1
> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

Este relatório é **gerado por `scripts/executar_testes.py`**: cada linha PASS/FAIL abaixo foi produzida por uma execução real. Nada foi marcado PASS sem executar. Para reproduzir: `python scripts/executar_testes.py --verificar`.

## 1. Resumo
- Testes automatizados: **90** — PASS **90**, FAIL **0**.
- Verificações de qualidade dos CSVs (Python, Q01–Q36): **36** — PASS **36**, FAIL **0**.
- Testes de detecção por injeção de erros: **36** — PASS **36**, FAIL **0**.
- Ambiente: Python 3.12.3, pandas 3.0.2, openpyxl 3.1.5, LibreOffice 24.2.7.2 420(Build:2), Linux x86_64.
- SHA-256 do XLSX (`excel/lobo_insights_industrial_v0_1.xlsx`): `cfdd4c65e4e53537fc4ff948d2651cb9ac1cfaeba3798e1919291d15e2068964`
- SHA-256 dos CSVs: `materiais_ficticios.csv` = `eca98597dc740956b4db95885ec54693def431a56501ea6d055fa54c1654231b`; `estoque_ficticio.csv` = `9f90cd52f208f9f863b2ababf7fed4735bbe893ef137efdc79f7506b3359a8d0`; `consumo_ficticio.csv` = `b5b8249731e5fcbce34b4345b4d1e576c568b33de3eba4dfca91f8bd0d94233c`

## 2. TESTADO × NÃO TESTADO
**TESTADO (executado nesta V0.1):** geração determinística dos CSVs; 36 verificações de qualidade em Python e o mesmo conjunto em fórmulas do Excel (recalculadas no LibreOffice); detecção de erros injetados (Python e Excel); estrutura do XLSX (abas, tabelas, nomes, validações, formatação condicional, ausência de vínculos externos); colunas calculadas; resumos por mês/categoria/centro/material; reconciliação dos 6 KPIs com um cálculo independente (módulo `csv` + `Decimal`); formato/codificação dos CSVs; varredura por regex de PII/credenciais; consistência da documentação com os dados.

**NÃO TESTADO (declarado, não assumido):**
- Abertura em Microsoft Excel real (Windows/Mac/Online). O arquivo foi lido com openpyxl e recalculado/renderizado com LibreOffice; comportamento no Excel não foi verificado.
- Excel com idioma pt-BR: as fórmulas são gravadas com nomes em inglês e o Excel deve traduzi-las ao abrir; isso não foi verificado.
- Semântica de célula vazia em `COUNTIF` (usada nos testes Q35 e derivados) foi verificada no LibreOffice; o Excel real pode diferir em casos extremos.
- Diferença de maiúsculas/minúsculas: `COUNTIFS` do Excel não diferencia caixa; o validador Python diferencia. Divergências só de caixa não foram testadas no Excel.
- Preenchimento automático das colunas calculadas ao adicionar linhas na tabela do Excel (não configurado nem testado).
- Tabelas dinâmicas (PivotTables): **não foram criadas** nesta versão.
- Dashboard/gráficos (V0.2), IA executiva (V0.3), classificador (V0.4), avaliação (V0.5), Power BI (V0.6), n8n (V0.7), GitHub (V1.0): fora do escopo, nada foi testado.
- Testes do PDF sobre IA e automação (número inventado, prioridade urgente, JSON inválido, entrada vazia): pertencem a versões futuras.
- A varredura de PII é heurística por regex: não prova a ausência absoluta de dado sensível; complementa a revisão humana.
- O conteúdo do ZIP final é conferido após o empacotamento, fora deste relatório (um arquivo não pode conter a prova do próprio ZIP).

## 3. Verificações de qualidade dos dados (Python, sobre os CSVs)
Resultado da execução de `scripts/validar_dados.py`. As mesmas 36 verificações existem como fórmulas na aba **Qualidade** do Excel.

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

## 4. Testes automatizados
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

## 5. Reconciliação dos 6 KPIs (Excel × CSVs)
O valor do CSV vem de `csv.DictReader` + `Decimal` (sem pandas e sem as fórmulas do Excel); o valor do Excel é o número calculado da aba Indicadores.

| KPI | Calculado dos CSVs | Excel (aba Indicadores) | Status |
|---|---|---|---|
| Unidades consumidas | 1647 | 1647 | PASS |
| Custo estimado consumido | 24606.20 | 24606.20 | PASS |
| Movimentações | 180 | 180 | PASS |
| Materiais distintos movimentados | 20 | 20 | PASS |
| Itens abaixo ou iguais ao mínimo | 5 | 5 | PASS |
| Itens críticos abaixo ou iguais ao mínimo | 2 | 2 | PASS |

## 6. Testes de detecção: as verificações realmente pegam erros?
Um dataset **com erros injetados de propósito** (cópia em memória; os CSVs oficiais não são alterados) é validado em Python e em um XLSX temporário com as mesmas fórmulas. "Esperado" foi definido à mão a partir do que foi injetado (incluindo efeitos colaterais, ex.: custo 0 no catálogo também diverge do estoque e do consumo).

| Teste | Erros injetados | Esperado | Python | Excel | PASS/FAIL |
|---|---|---|---|---|---|
| Q01 | linha MAT016 duplicada no catálogo | 2 | 2 | 2 | PASS |
| Q02 | linha MAT016 duplicada no estoque | 2 | 2 | 2 | PASS |
| Q03 | Movimento_ID repetido em duas linhas | 2 | 2 | 2 | PASS |
| Q04 | catálogo MAT010 com Fornecedor vazio | 1 | 1 | 1 | PASS |
| Q05 | estoque MAT020 com Fornecedor vazio | 1 | 1 | 1 | PASS |
| Q06 | consumo com Tipo vazio | 1 | 1 | 1 | PASS |
| Q07 | quantidades 0, -3 e 2,5 | 3 | 3 | 3 | PASS |
| Q08 | catálogo MAT019 com custo 0 | 1 | 1 | 1 | PASS |
| Q09 | estoque MAT007 com custo -1 | 1 | 1 | 1 | PASS |
| Q10 | consumo com custo 0 | 1 | 1 | 1 | PASS |
| Q11 | estoque com Material_ID inexistente (MAT999) | 1 | 1 | 1 | PASS |
| Q12 | consumo com MAT999 | 1 | 1 | 1 | PASS |
| Q13 | (nenhum erro injetado) | 0 | 0 | 0 | PASS |
| Q14 | nome do estoque MAT008 alterado | 1 | 1 | 1 | PASS |
| Q15 | nome do material alterado no consumo | 1 | 1 | 1 | PASS |
| Q16 | estoque MAT018 diverge do catálogo (categoria); categoria do estoque MAT009 = EPI (válida, mas diverge) | 2 | 2 | 2 | PASS |
| Q17 | consumo de MAT018 diverge do catálogo (categoria); categoria 'Papelaria' no consumo diverge do catálogo | 4 | 4 | 4 | PASS |
| Q18 | estoque MAT006 diverge do catálogo (unidade); unidade do estoque MAT015 = UN (válida, mas diverge) | 2 | 2 | 2 | PASS |
| Q19 | consumo de MAT006 diverge do catálogo (unidade); unidade do consumo = PAR (válida, mas diverge) | 5 | 5 | 5 | PASS |
| Q20 | estoque MAT019 diverge do catálogo (custo 0); estoque MAT007 diverge do catálogo (custo -1) | 2 | 2 | 2 | PASS |
| Q21 | consumo de MAT019 diverge do catálogo (custo 0); consumo com custo 0 diverge do catálogo; custo do consumo = 99,99 (diverge) | 3 | 3 | 3 | PASS |
| Q22 | estoque MAT014 diverge do catálogo (criticidade); criticidade do estoque MAT001 = Alta (diverge) | 2 | 2 | 2 | PASS |
| Q23 | estoque MAT010 diverge do catálogo (fornecedor vazio); estoque MAT009 diverge do catálogo (fornecedor); estoque MAT020 diverge do catálogo (fornecedor vazio) | 3 | 3 | 3 | PASS |
| Q24 | data inexistente 2026-13-45; data fora do período 2026-09-15 | 2 | 2 | 2 | PASS |
| Q25 | (nenhum erro injetado) | 0 | 0 | 0 | PASS |
| Q26 | estoque MAT011 = -2 | 1 | 1 | 1 | PASS |
| Q27 | mínimo MAT012 = -5 | 1 | 1 | 1 | PASS |
| Q28 | lead time 0 (MAT013) e 2,5 (MAT017) | 2 | 2 | 2 | PASS |
| Q29 | catálogo MAT018 com categoria inválida | 1 | 1 | 1 | PASS |
| Q30 | (nenhum erro injetado) | 0 | 0 | 0 | PASS |
| Q31 | consumo com categoria 'Papelaria' | 1 | 1 | 1 | PASS |
| Q32 | catálogo MAT006 com unidade inválida | 1 | 1 | 1 | PASS |
| Q33 | catálogo MAT014 com criticidade inválida | 1 | 1 | 1 | PASS |
| Q34 | centro 'FUNDICAO' | 1 | 1 | 1 | PASS |
| Q35 | Tipo vazio também é tipo inválido; tipo 'Devolucao' | 2 | 2 | 2 | PASS |
| Q36 | Fornecedor vazio também fica fora do padrão fictício; catálogo MAT009 com fornecedor fora do padrão | 2 | 2 | 2 | PASS |

## 7. Inspeção visual do Excel
O XLSX final foi renderizado com o LibreOffice (conversão para PDF de 11 páginas, depois PNG) e as imagens foram **inspecionadas visualmente**:
- Build final (SHA-256 acima): LEIA_ME (p. 1), Estoque (p. 3), Qualidade (p. 8) e Indicadores (p. 9).
- Build imediatamente anterior, que difere do final apenas por alinhamento de células (centralização de colunas numéricas): Materiais (p. 2), primeira página de Consumo (p. 4), Indicadores segunda página (p. 10) e Parametros (p. 11).
- **Não visualizadas:** páginas 5 a 7 (continuação da tabela Consumo).
- Observado: cabeçalhos azuis/verdes, formatação condicional de REPOR/ATENÇÃO, coluna Status PASS em verde, blocos de resumo e conferências legíveis.
- **Não realizada** inspeção visual no Microsoft Excel real; a renderização é do LibreOffice e pode diferir do Excel.


## 8. Registro de problemas e observações da execução
- **Bases geradas:** a primeira geração dos CSVs passou nas 36 verificações sem violação de regra; portanto não houve correção de dados a registrar. A semente (20260601) foi fixada antes de ver o resultado e não foi ajustada.
- **Observação sobre o dado:** agosto tem menos movimentações (40) que junho (65) e julho (75). Resulta do sorteio aleatório; a distribuição desigual é requisito do enunciado e nenhuma conclusão foi forçada. Ver `LIMITATIONS.md`, item 2.
- **Ajustes após inspeção visual do Excel:** centralização de colunas numéricas (Estoque, Materiais, Consumo), alinhamento vertical dos KPIs e remoção do valor fixo "25%" do texto do LEIA_ME (a margem é um parâmetro editável).
- **Inconsistências do PDF:** 10 pontos ambíguos documentados em `V0_1_NOTES.md`, seção 4, sem alteração silenciosa.

